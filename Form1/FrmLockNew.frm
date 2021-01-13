VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FrmLockNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lock Account New"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "List Custid Preview"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   10560
      TabIndex        =   78
      Top             =   0
      Width           =   3975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   7320
         Width           =   615
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   6795
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   11986
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label11 
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   82
         Top             =   7320
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5355
      Left            =   60
      TabIndex        =   37
      Top             =   60
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   9446
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Lock Status Account"
      TabPicture(0)   =   "FrmLockNew.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Lock Markup"
      TabPicture(1)   =   "FrmLockNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "User Lock"
      TabPicture(2)   =   "FrmLockNew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdCekAll"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LvUser"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command2"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Lock Berdasar Payment"
      TabPicture(3)   =   "FrmLockNew.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "FrameBlokPayment"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "CeKBlokPayment"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lainnya"
      TabPicture(4)   =   "FrmLockNew.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chk_ptp_payment"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Line1"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.CheckBox chk_ptp_payment 
         Caption         =   "PTP sinkronisasi dengan pembayaran bulan ini"
         Height          =   255
         Left            =   -74640
         TabIndex        =   77
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox CeKBlokPayment 
         Caption         =   "Tampilkan data hanya dengan payment:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   71
         Top             =   780
         Width           =   3795
      End
      Begin VB.Frame FrameBlokPayment 
         Enabled         =   0   'False
         Height          =   3885
         Left            =   300
         TabIndex        =   72
         Top             =   840
         Width           =   9765
         Begin VB.OptionButton OptPayment1Bln 
            Caption         =   "1 bulan lalu"
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
            Left            =   240
            TabIndex        =   76
            Top             =   390
            Width           =   2205
         End
         Begin VB.OptionButton OptPayment2Bln 
            Caption         =   "2 bulan lalu"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   75
            Top             =   750
            Width           =   2175
         End
         Begin VB.OptionButton OptPayment3Bln 
            Caption         =   "3 bulan lalu"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   74
            Top             =   1140
            Width           =   2925
         End
         Begin VB.OptionButton OptLbh3Bln 
            Caption         =   "< 3 bulan lalu"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   73
            Top             =   1500
            Width           =   2265
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&UnCek All"
         Height          =   375
         Left            =   -73560
         TabIndex        =   70
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cek Berdasarkan SPV Code"
         Height          =   615
         Left            =   -72300
         TabIndex        =   65
         Top             =   4380
         Width           =   5955
         Begin VB.CommandButton CmdUncekSPV 
            Caption         =   "UnCek SPV"
            Height          =   315
            Left            =   4440
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton CmdCekSpv 
            Caption         =   "Cek SPV"
            Height          =   315
            Left            =   3120
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox CmbSPV 
            Height          =   315
            Left            =   1140
            TabIndex        =   67
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label Label10 
            Caption         =   "SPV Code"
            Height          =   255
            Left            =   180
            TabIndex        =   66
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "&Cek All"
         Height          =   375
         Left            =   -74820
         TabIndex        =   64
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   4425
         Left            =   -74880
         TabIndex        =   50
         Top             =   480
         Width           =   10185
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   4560
            TabIndex        =   57
            Top             =   2640
            Width           =   825
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   4560
            TabIndex        =   56
            Top             =   2220
            Width           =   825
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   4560
            TabIndex        =   55
            Top             =   1680
            Width           =   825
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   4560
            TabIndex        =   54
            Top             =   1260
            Width           =   825
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   375
            Left            =   4560
            TabIndex        =   51
            Top             =   840
            Width           =   825
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   3315
            Left            =   60
            TabIndex        =   58
            Top             =   690
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   5847
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   3315
            Left            =   5580
            TabIndex        =   59
            Top             =   690
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   5847
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
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
         Begin VB.CheckBox chksingle 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Single"
            Height          =   345
            Left            =   780
            TabIndex        =   53
            Top             =   1260
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkmultiple 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multiple"
            Height          =   345
            Left            =   600
            TabIndex        =   52
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Mark Up"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   420
            Width           =   1185
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Lead Markup"
            Height          =   255
            Left            =   5580
            TabIndex        =   60
            Top             =   420
            Width           =   2685
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4395
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   9975
         Begin VB.CheckBox Check1 
            Caption         =   "S P - Settled Payment"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   49
            Top             =   1290
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RP - Refuse Payment"
            Height          =   255
            Index           =   7
            Left            =   135
            TabIndex        =   48
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VL-Valid"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "OS - On Process"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "P T P - Promise To Pay"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   1680
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "B P - Broken Promise"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   44
            Top             =   600
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "P O P - Progress Of Payment"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   43
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Blank Data"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   42
            Top             =   1650
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PR-Prospect"
            Height          =   255
            Index           =   9
            Left            =   5700
            TabIndex        =   41
            Top             =   630
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ON-On Nego"
            Height          =   255
            Index           =   10
            Left            =   5700
            TabIndex        =   40
            Top             =   930
            Width           =   1245
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SK-Skip"
            Height          =   255
            Index           =   11
            Left            =   5670
            TabIndex        =   39
            Top             =   1260
            Width           =   1245
         End
      End
      Begin MSComctlLib.ListView LvUser 
         Height          =   3915
         Left            =   -74820
         TabIndex        =   62
         Top             =   420
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6906
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   -74400
         TabIndex        =   63
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   -74760
         X2              =   -65040
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame FrameEntry 
      BackColor       =   &H00E0E0E0&
      Height          =   2025
      Left            =   60
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CheckBox chknewentry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   11
         Top             =   300
         Width           =   1725
      End
      Begin VB.CheckBox chkreguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   690
         Width           =   1485
      End
      Begin VB.CheckBox chkswap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   9
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkcurrent 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   8
         Top             =   1530
         Width           =   1755
      End
      Begin VB.OptionButton OptSwap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox CmbSwap 
         Height          =   315
         ItemData        =   "FrmLockNew.frx":008C
         Left            =   3660
         List            =   "FrmLockNew.frx":0093
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox CmbReguler 
         Height          =   315
         ItemData        =   "FrmLockNew.frx":009C
         Left            =   3660
         List            =   "FrmLockNew.frx":00A3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptReguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.ComboBox CmbNewEntry 
         Height          =   315
         ItemData        =   "FrmLockNew.frx":00AC
         Left            =   3660
         List            =   "FrmLockNew.frx":00B3
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptNewEntry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New entry "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5355
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   12
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.CheckBox CheckEntry 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tampilkan data entry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   780
      TabIndex        =   0
      Top             =   3900
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Timer Timer_stopwatch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7770
      Top             =   3660
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CommandButton CmdAddUser 
         Caption         =   "&User.."
         Height          =   375
         Left            =   7860
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox CHKACCOUNT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS COMPLETE"
         Height          =   285
         Left            =   5970
         TabIndex        =   19
         Top             =   570
         Width           =   1845
      End
      Begin VB.CheckBox CHKLUNASPENDING 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS PENDING"
         Height          =   465
         Left            =   5970
         TabIndex        =   18
         Top             =   60
         Width           =   1725
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   17
         Top             =   600
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   16
         Text            =   "XXX"
         Top             =   600
         Width           =   1335
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "All TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih TeleCollection"
         Value           =   -1
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih SPV Name"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TeleCollection Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   7260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Add Schedule"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   1980
      TabIndex        =   25
      Top             =   7260
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Release"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   26
      Top             =   7260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "E&xit"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   3
      Left            =   3060
      TabIndex        =   27
      Top             =   7260
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&SHUT"
      ButtonStyle     =   3
   End
   Begin TDBDate6Ctl.TDBDate StartDate 
      Height          =   315
      Left            =   1515
      TabIndex        =   28
      Top             =   5565
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "FrmLockNew.frx":00BC
      Caption         =   "FrmLockNew.frx":01D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmLockNew.frx":0240
      Keys            =   "FrmLockNew.frx":025E
      Spin            =   "FrmLockNew.frx":02BC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime StartTime 
      Height          =   315
      Left            =   3090
      TabIndex        =   29
      Top             =   5565
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "FrmLockNew.frx":02E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmLockNew.frx":0350
      Spin            =   "FrmLockNew.frx":03A0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate EndDate 
      Height          =   315
      Left            =   1515
      TabIndex        =   30
      Top             =   5985
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "FrmLockNew.frx":03C8
      Caption         =   "FrmLockNew.frx":04E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmLockNew.frx":054C
      Keys            =   "FrmLockNew.frx":056A
      Spin            =   "FrmLockNew.frx":05C8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime EndTime 
      Height          =   315
      Left            =   3090
      TabIndex        =   31
      Top             =   5985
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "FrmLockNew.frx":05F0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmLockNew.frx":065C
      Spin            =   "FrmLockNew.frx":06AC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
      Value           =   1.02960316199441E-317
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   4
      Left            =   8520
      TabIndex        =   79
      Top             =   7260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Preview >>"
      ButtonStyle     =   3
   End
   Begin VB.Label Label8 
      Caption         =   "End Lock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   35
      Top             =   5985
      Width           =   1170
   End
   Begin VB.Label Label7 
      Caption         =   "Start Lock:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   150
      TabIndex        =   34
      Top             =   5565
      Width           =   1170
   End
   Begin VB.Label LblWaktu 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Label Waktu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6060
      TabIndex        =   33
      Top             =   5940
      Width           =   4380
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Waktu Server Sekarang:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6030
      TabIndex        =   32
      Top             =   5550
      Width           =   4380
   End
End
Attribute VB_Name = "FrmLockNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalTenthDetik, TotalDetik, TenthDetik, Detik, Menit, JAM As Integer
Dim Jam1 As String

Dim CMDSQL As String
Dim StsVl As String
Dim StsOS As String
Dim StsON As String
Dim StsSK As String
Dim StsPR As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsRP As String
Dim StsOP As String
Dim StsFresh As String
Dim Stsblank As String
Dim Stsuncontact As String
Dim spv As Boolean
'@@ 140710 Tambahan buat blok entry yang diambil dari field entry_date dan pay_dt di mgm
Dim BlokEntry As String
'@@ 061110 Blok data automatic dengan waktu
Dim StringBlokTimer As String
Dim StatusLocked As String
'@@ 18-11-10 Perbaikan dari blok data entry
Dim StsNewEntry As String
Dim StsReguler As String
Dim StsSwap As String
Dim StsCurrent As String
Dim CekValidLock As Boolean

'@@ 15 Agustus 2011 Buat blok payment
Dim StsPayment1Bln As String
Dim StsPayment2Bln As String
Dim StsPayment3Bln As String
Dim StsPaymentLbh3Bln As String

Private Sub CeKBlokPayment_Click()
    If FrameBlokPayment.Enabled = False Then
        FrameBlokPayment.Enabled = True
    Else
        FrameBlokPayment.Enabled = False
    End If
    
    If OptPayment1Bln.Value = False Then
        OptPayment1Bln.Value = True
    Else
        OptPayment1Bln.Value = False
    End If
    
    OptPayment2Bln.Value = False
    OptPayment3Bln.Value = False
    OptLbh3Bln.Value = False
End Sub


Private Sub CmdAddUser_Click()
    FrmDaftarUserLock.Show vbModal
End Sub

Private Sub CmdCekAll_Click()
    Dim F As Integer
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For i = 1 To LvUser.ListItems.Count
        LvUser.ListItems(i).Checked = True
    Next i
End Sub

Private Sub CmdCekSpv_Click()
    Dim F As Integer
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Data User tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For F = 1 To LvUser.ListItems.Count
        If LvUser.ListItems(F).SubItems(2) = UCase(Trim(CmbSpv.Text)) Then
            LvUser.ListItems(F).Checked = True
        End If
    Next F
End Sub

Private Sub CmdUncekSPV_Click()
    Dim F As Integer
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Data User tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For F = 1 To LvUser.ListItems.Count
        If LvUser.ListItems(F).SubItems(2) = UCase(Trim(CmbSpv.Text)) Then
            LvUser.ListItems(F).Checked = False
        End If
    Next F
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
         KeyAscii = 0
        Case 1
         KeyAscii = 0
    End Select
End Sub

Private Sub Command2_Click()
    CMDSQL = "('a',"
    For i = 1 To LvUser.ListItems.Count
       
        If LvUser.ListItems(i).Checked = True Then
            CMDSQL = CMDSQL + Trim(LvUser.ListItems(i).Text) + ","
        End If
    Next i
    MsgBox CMDSQL
End Sub



Private Sub Command3_Click()
    Dim F As Integer
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For i = 1 To LvUser.ListItems.Count
        LvUser.ListItems(i).Checked = False
    Next i
End Sub

Private Sub Form_Activate()
    'SSOption1(2).Value = True
    'SSOption1(2).Enabled = False
    'Combo1(0).Text = Replace(Trim(mdiform1.txtusername.text), "TL", "SPV")
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Batch", 10 * 1200
    ListView2.ColumnHeaders.ADD 1, , "Batch", 20 * 1200
    
    ListView3.ColumnHeaders.ADD 1, , "CUST ID", 20 * 120
End Sub
Private Sub CheckEntry_Click()
    If FrameEntry.Enabled = False Then
        FrameEntry.Enabled = True
        OptNewEntry.Value = True
    Else
        FrameEntry.Enabled = False
        OptNewEntry.Value = False
        OptReguler.Value = False
        OptSwap.Value = False
    End If
End Sub


Private Sub chkmultiple_Click()
    If chkmultiple.Value = vbChecked Then
        chksingle.Value = vbUnchecked
    End If
End Sub

Private Sub chksingle_Click()
    If chksingle.Value = vbChecked Then
        chkmultiple.Value = vbUnchecked
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
Dim i As Integer

Select Case Index
Case 1
    If ListView2.ListItems.Count <> 0 Then
            Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
    End If
Case 3
    For i = 1 To ListView2.ListItems.Count
                Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
                ListView2.ListItems.Remove ListView2.SelectedItem.Index
    Next
Case 0
'    If ListView1.ListItems.Count <> 0 Then
'        Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
'        ListView1.ListItems.Remove ListView1.SelectedItem.Index
'    End If
If ListView1.ListItems.Count <> 0 Then
ulang:
    For i = 1 To ListView1.ListItems.Count - 1
        If ListView1.ListItems(i).Checked = True Then
            Set lList = ListView2.ListItems.ADD(, , ListView1.ListItems(i).Text)
            ListView1.ListItems.Remove ListView1.ListItems(i).Index
            GoTo ulang
        End If
    Next i
End If

Case 2
    For i = 1 To ListView1.ListItems.Count
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
                   
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
    Next
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    If spv = False Then
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from mandiri.SPVTBL where SPVCODE='" + Combo1(0) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
Case 1
    If spv = False Then
        Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from mandiri.SPVTBL where SPVNAME='" + Combo1(1) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
    
 End Select
 Set m_data = Nothing
 Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    '@@ 19-11-10 tambahan kasih konfirmasi dan cek dulu datanya
    Dim a As String
    Dim SqlCek As String
    Dim M_objrs As ADODB.Recordset
    
    a = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
ulang:
        For i = 1 To ListView1.ListItems.Count - 1
            If ListView1.ListItems(i).Checked = True Then
                SqlCek = "select lockmarkup from mandiri.usertbl where lockmarkup like '%"
                SqlCek = SqlCek + Trim(ListView1.SelectedItem.Text) + "%'"
                
                Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open SqlCek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_objrs.RecordCount <> 0 Then
                    MsgBox "Data tidak dapat dihapus! Karena data masih dalam proses blok data!", vbOKOnly + vbExclamation, "Peringatan"
                    Set M_objrs = Nothing
                    GoTo lompat
                Else
                    M_OBJCONN.Execute "UPDATE mandiri.MGM SET EXCLUDE =NULL WHERE EXCLUDE='" + Trim(ListView1.ListItems(i).Text) + "'"
                    'getMarkup
                    ListView1.ListItems.Remove (ListView1.ListItems(i).Index)
                    'Exit Sub
                    GoTo ulang
                    MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
                End If
            End If
lompat:
        Next i
        
    End If
    Set M_objrs = Nothing
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim m_data As New CLS_FRMSEARCH
    Dim m_waktuserver As ADODB.Recordset
    Dim SqlWaktu As String
    
    
    Set M_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "")
        While Not M_objrs.EOF
            Combo1(0).AddItem cnull(M_objrs("USERID"))
            Combo1(1).AddItem M_objrs("AGENT")
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
    
    'Jika yang login tl maka blok data ke semua agent di nonaktifkan
    '@@ 01-03-2011, di remarks nih
'    If Left(Trim(mdiform1.txtusername.text), 2) = "TL" Then
'        SSOption1(0).Enabled = False
'        SSOption1(1).Enabled = True
'        SSOption1(1).Value = True
'    Else
'        SSOption1(0).Value = True
'    End If
        
        
    spv = False
    header
    CmbNewEntry.Text = "< 2"
    CmbReguler.Text = "< 2"
    CmbSwap.Text = "> 2"
    getMarkup
    
    'Ambil waktu server
    SqlWaktu = "select now()"
    Set m_waktuserver = New ADODB.Recordset
    m_waktuserver.CursorLocation = adUseClient
    m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        JAM = Val(Format(m_waktuserver(0), "hh"))
        Menit = Val(Format(m_waktuserver(0), "nn"))
        Detik = Val(Format(m_waktuserver(0), "ss"))
    Set m_waktuserver = Nothing
    LblWaktu.Caption = JAM & ":" & Menit
   ' Timer_stopwatch.Enabled = True
   
   '@@ 01-03-2011 membuat daftar user
   Call headerUser
   Call IsiUser
   Call IsiSpv
End Sub





Private Sub LvUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvUser.SortKey = ColumnHeader.Index - 1
    LvUser.Sorted = True
End Sub

Private Sub SSCommand1_Click(Index As Integer)
    Dim M_objrs As New ADODB.Recordset
    Dim sStrsql As String
    Dim mwhere As String
    Dim isi_ceklistuser As Boolean
    
    Select Case Index
    Case 0
       
        '@@ 061110 - akhir blok dengan timer -
            StatusLocked = ""
            
    
        
    If CHKLUNASPENDING.Value = vbChecked And CHKACCOUNT.Value = vbChecked Then
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
            sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
           'mwhere = " WHERE SPVCODE LIKE '%SPV%'"
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            mwhere = "Where spvcode='" + Combo1(0).Text + "'"
            STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
            STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
            STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
            'M_OBJCONN.Execute STRSQL
            
            '@@ 061110 - awal blok dengan timer -
              StringBlokTimer = StringBlokTimer + STRSQL + " | "
              StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
            '@@ 061110 - akhir blok dengan timer -
    
            
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
                Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
                While Not M_objrs.EOF
                    sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) "
                    sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                    mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                    'M_OBJCONN.Execute STRSQL
                    
                    '@@ 061110 - awal blok dengan timer -
                    StringBlokTimer = StringBlokTimer + STRSQL + " | "
                    '@@ 061110 - akhir blok dengan timer -
                    
                    M_objrs.MoveNext
                Wend
                '@@ 061110
                StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
                '@@ 061110
                Set M_objrs = Nothing
        Else
        Exit Sub
        End If
        
        'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
      ElseIf CHKLUNASPENDING.Value = vbChecked Then
            If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
              sStrsql = " agent in (@LUNAS PENDING@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
                mwhere = " WHERE SPVCODE LIKE '%SPV%'"
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
                mwhere = "Where spvcode='" + Combo1(0).Text + "'"
                 STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                 STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                 STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                'M_OBJCONN.Execute STRSQL
                
             '@@ 061110 - awal blok dengan timer -
              StringBlokTimer = StringBlokTimer + STRSQL + " | "
             '@@ 061110 - akhir blok dengan timer -
             
             '@@ 061110
                StatusLocked = StatusLocked + " LUNAS PENDING- "
              '@@ 061110
                
            ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
                Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
                While Not M_objrs.EOF
                    sStrsql = " agent in (@LUNAS PENDING@) "
                    sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                    mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                    'M_OBJCONN.Execute STRSQL
                    
                    '@@ 061110 - awal blok dengan timer -
                    StringBlokTimer = StringBlokTimer + STRSQL + " | "
                    '@@ 061110 - akhir blok dengan timer -
                    
                    M_objrs.MoveNext
                Wend
                '@@ 061110
                StatusLocked = StatusLocked + " LUNAS PENDING-"
                '@@ 061110
                Set M_objrs = Nothing
        Else
            Exit Sub
            
            End If
            
            'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
            Exit Sub
    ElseIf CHKACCOUNT.Value = vbChecked Then
               sStrsql = " agent in (@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
                mwhere = " WHERE SPVCODE LIKE '%SPV%'"
            If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
                mwhere = "Where spvcode='" + Combo1(0).Text + "'"
                STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                '@@ 061110
                StatusLocked = StatusLocked + " LUNAS COMPLETE -"
                '@@ 061110
                
            ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
                Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
                While Not M_objrs.EOF
                    sStrsql = " agent in (@LUNAS COMPLETE@) "
                    sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM mandiri.USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                    mwhere = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + mwhere
                    'M_OBJCONN.Execute STRSQL
                    
                    '@@ 061110 - awal blok dengan timer -
                    StringBlokTimer = StringBlokTimer + STRSQL + " | "
                    '@@ 061110 - akhir blok dengan timer -
                    
                    M_objrs.MoveNext
                Wend
                Set M_objrs = Nothing
                
                '@@ 061110
                StatusLocked = StatusLocked + " LUNAS COMPLETE -"
                '@@ 061110
                
        Else
            Exit Sub
            
            End If
            
                'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
            Exit Sub
    End If
    
    
            
            If SSOption1(0).Value = False And SSOption1(1).Value = False And SSOption1(2).Value = False Then
                MsgBox "Select DCR Name To Proccess OR All"
             Else
                    If SSOption1(0).Value Then
                        Call ceksts
                        STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where usertype='1'"
                        M_OBJCONN.Execute (STRSQL)
                        
                        
                        STRSQL = "UPDATE mandiri.usertbl SET f_flagrender=1, lockdarispv ='"
                        STRSQL = STRSQL + getblock + "',lock_entry_lpd='"
                        STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                        STRSQL = STRSQL + " Where usertype='1'"
                        'M_OBJCONN.Execute (STRSQL)
                        
                        '@@ 061110 - awal blok dengan timer -
                        StringBlokTimer = StringBlokTimer + STRSQL + " | "
                        '@@ 061110 - akhir blok dengan timer -
                        
                        
                            If ListView2.ListItems.Count <> 0 Then
                                HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                                sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                                sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                                sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                                sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ")) AS C "
                                sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1' "
                                'M_OBJCONN.Execute (sStrsql)
                                
                                '@@ 061110 - awal blok dengan timer -
                                StringBlokTimer = StringBlokTimer + sStrsql + " | "
                                '@@ 061110 - akhir blok dengan timer -
                                '@@ 061110
                                StatusLocked = StatusLocked + HLSMARKUP + "-"
                                '@@ 061110
                            End If
                        'MsgBox "Proccess to All DCR Name Done.....!"
                      End If
            
                    If SSOption1(1).Value Then
                        If Combo1(0).Text = "" Then
                            MsgBox "Select DCR Name To Proccess..!"
                            Combo1(0).SetFocus
                             
                        Else
                            Call ceksts
                            '@@ 01-03-2011 Di remarks dulu
    '                        StrSql = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
    '                        StrSql = StrSql + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
    '                        StrSql = StrSql + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where userid='" + Combo1(0).Text + "'"
    '                        M_OBJCONN.Execute (StrSql)
                            
    '                        StrSql = "UPDATE usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
    '                        StrSql = StrSql + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
    '                        StrSql = StrSql + mdiform1.txtusername.text + "-" + Format(Now, "yyyy-mm-dd") + "' "
    '                        StrSql = StrSql + "Where userid='" + Trim(Combo1(0).Text) + "'"
    '                        'M_OBJCONN.Execute STRSQL
    '
    '                        '@@ 061110 - awal blok dengan timer -
    '                        StringBlokTimer = StringBlokTimer + StrSql + " | "
    '                        '@@ 061110 - akhir blok dengan timer -
    '
    '                        If ListView2.ListItems.Count <> 0 Then
    '                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
    '                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
    '                            sStrsql = sStrsql + mdiform1.txtusername.text + "-" + Format(Now, "yyyy-mm-dd") + "',"
    '                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
    '                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT='" + Combo1(0).Text + "') AS C "
    '                            sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1'"
    '                            'M_OBJCONN.Execute (sStrsql)
    '
    '                            '@@ 061110 - awal blok dengan timer -
    '                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
    '                            '@@ 061110 - akhir blok dengan timer -
    '
    '                            '@@ 061110
    '                            StatusLocked = StatusLocked + HLSMARKUP + "-"
    '                            '@@ 061110
    '
    '                        End If
                            
                        
                            'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        End If
                    Else
                        If SSOption1(2).Value = True Then
                            If Combo1(0).Text = "" Then
                                MsgBox "Select SPV Name To Proccess..!"
                            Else
                                Call ceksts
                            STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL,F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                            STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                            STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lockmarkup=NULL,lockdarispv=null,lock_entry_lpd=null Where spvcode='" + Combo1(0).Text + "'"
                            M_OBJCONN.Execute (STRSQL)
                                
                           'If CHKLUNASPENDING.Value = vbChecked Then
                            '    STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispvbuattl ='" + getblock + "',lock_entry_lpd='"
                             '   STRSQL = STRSQL + BlokEntry + "',fromaccount ='" + cboaccount.Text + "' Where spvcode='" + Combo1(0).Text + "'"
                              '  M_OBJCONN.Execute STRSQL
                           'End If
                           
                           'If CHKACCOUNT.Value = vbChecked Then
                                STRSQL = "UPDATE mandiri.usertbl SET f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                                STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                                STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                                STRSQL = STRSQL + "Where spvcode='" + Trim(Combo1(0).Text) + "' and usertype='1'"
                                'M_OBJCONN.Execute STRSQL
                           ' End If
                           
                                '@@ 061110 - awal blok dengan timer -
                                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                                '@@ 061110 - akhir blok dengan timer -
                            
                                
                            If ListView2.ListItems.Count <> 0 Then
                                HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                                sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                                sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                                sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                                sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT IN (SELECT USERID FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "')) AS C "
                                sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1' "
                                'M_OBJCONN.Execute (sStrsql)
                                
                                '@@ 061110 - awal blok dengan timer -
                                StringBlokTimer = StringBlokTimer + sStrsql + " | "
                                '@@ 061110 - akhir blok dengan timer -
                                
                                '@@ 061110
                                StatusLocked = StatusLocked + HLSMARKUP + "-"
                                '@@ 061110
                            End If
                            
                                
                                'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                            End If
                        End If
                 End If
            End If
            
            'Cek validitas status yang di lock
            Call CekValidNullLock
            If CekValidLock = False Then
                Exit Sub
            End If
            
            
            ' # CEK validasi list user
            isi_ceklistuser = False
            
            For i = 1 To LvUser.ListItems.Count
                If LvUser.ListItems(i).Checked = True Then
                    isi_ceklistuser = True
                End If
            Next i
            
            If Not isi_ceklistuser Then
                MsgBox "Anda belum memilih user yang akan di lock!!", vbOKOnly + vbInformation, "INFO"
                Exit Sub
            End If
            ' # END CEK
          
           ' CEK pilih waktu
            If StartDate.ValueIsNull Then
                MsgBox "Anda belum memilih Start Date!!!", vbCritical + vbOKOnly, "INFO"
                Exit Sub
            End If
            
            If EndDate.ValueIsNull Then
                MsgBox "Anda belum memilih End Date!!!", vbCritical + vbOKOnly, "INFO"
                Exit Sub
            End If
            ' # END CEK
           
            Dim WaktuAwal As Date
            Dim WaktuAkhir As Date
            Dim WaktuServer As Date
            Dim m_ObjrsWktServer As ADODB.Recordset
                        
            Set m_ObjrsWktServer = New ADODB.Recordset
            m_ObjrsWktServer.CursorLocation = adUseClient
            m_ObjrsWktServer.Open "select now() as waktu ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            WaktuServer = Format(m_ObjrsWktServer(0), "mm-dd-yyyy hh:mm")
            Set m_ObjrsWktServer = Nothing
                                    
            WaktuAwal = Format(StartDate.Value, "mm-dd-yyyy") & " " & StartTime.Value
            WaktuAkhir = Format(EndDate.Value, "mm-dd-yyyy") & " " & EndTime.Value
                        
            '@@ 19-11-10 buat validasi pembuatan schedule
            'cek dulu tanggalnya udah diisi apa belum??
            If IsNull(StartDate.Value) Or IsNull(EndDate.Value) Or IsNull(StartTime.Value) Or IsNull(EndTime.Value) Then
                MsgBox "Start date dan End date, tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
                StatusLocked = ""
                Exit Sub
            End If
                        
            'Start Lock tidak boleh lebih kecil dari waktu server
            If WaktuAwal < WaktuServer Then
                MsgBox "Start lock tidak boleh lebih kecil dari waktu server!", vbOKOnly + vbExclamation, "Peringatan"
                StatusLocked = ""
                Exit Sub
            End If
                                            
            'Waktu Awal tidak boleh sama dengan waktu akhir
            If WaktuAwal = WaktuAkhir Then
                MsgBox "Start Lock tidak boleh sama dengan End Lock!", vbOKOnly + vbExclamation, "Peringatan"
                StatusLocked = ""
                Exit Sub
            End If
                        
            'Cek waktu awal tidak boleh lebih besar dari wktu akhir
            If WaktuAwal > WaktuAkhir Then
                MsgBox "Waktu awal tidak boleh lebih besar dari waktu akhir!", vbOKOnly + vbExclamation
                StatusLocked = ""
                Exit Sub
            End If
           
           
                '@@ 01-03-2011 Blok Data berdasarkan pilihan agent
                Dim z As Integer
                Call ClearData
                For z = 1 To LvUser.ListItems.Count
                    If LvUser.ListItems(z).Checked = True Then
                        
                         '@@ 061110 - awal blok dengan timer -
                         StringBlokTimer = "AWAL | "
                    
                        'Lock entry_lpd , lock dari spsv
    '                    Strsql = "UPDATE usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
    '                    Strsql = Strsql + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
    '                    Strsql = Strsql + mdiform1.txtusername.text + "-" + Format(Now, "yyyy-mm-dd") + "', "
    '                    Strsql = Strsql + " lockpayment='" + GetBlokPayment + "' "
    '                    Strsql = Strsql + " Where userid='" + Trim(LvUser.ListItems(z).Text) + "'"
                        STRSQL = "UPDATE mandiri.usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_ptp_payment='" + getBlokPTPNoPayment + "',lock_entry_lpd='"
                        STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "', "
                        STRSQL = STRSQL + " lockpayment='" + GetBlokPayment + "' "
                        STRSQL = STRSQL + " Where userid='" + Trim(LvUser.ListItems(z).Text) + "'"
                            
                        '@@ 061110 - awal blok dengan timer -
                        StringBlokTimer = StringBlokTimer + STRSQL + " | "
                        '@@ 061110 - akhir blok dengan timer -
                    
                        'Ini buat yang markup
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE mandiri.USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM mandiri.MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT='" + Trim(LvUser.ListItems(z).Text) + "') AS C "
                            sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1'"
                                    
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                                    
                            '@@ 061110
                            'StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                        End If
                        
                        
                        '@@ 061110 - awal blok dengan timer -
                        StringBlokTimer = Replace(StringBlokTimer + " AKHIR ", "'", "$")
                        
                        Dim CmdTimer As String
                        Dim AccLock As String
                        
                        If SSOption1(0).Value = True Then
                            AccLock = "ALL"
                        Else
                           AccLock = Trim(LvUser.ListItems(z).Text)
                        End If
                        
                        
                        
                        'Cek apakah dalam tanggal waktu yang sama, ada lock account yang sudah dibuat dlm antrian
                        Dim CekSch As String
                        Dim M_Objrs_CekSch As ADODB.Recordset
                       
                       'cek apakah dia sebagai teamleader or spv or admin
            '           If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
            '                 'Ini jika yang login TL
            '                 CekSch = "select * from tbltemplockacc where date(start_lock)='"
            '                 CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and f_locked isnull and (lock_by='"
            '                 CekSch = CekSch + Trim(mdiform1.txtusername.text) + "' or lock_by in ('SEPTIAN','JOKO','WULAN','ADMINISTRATOR','ADMIN','SUPERVISOR'))"
            '                 CekSch = CekSch + " and account_lock in ('" + Trim(AccLock) + "','ALL')"
            '           Else
            '                'Ini jika yang login SPV or Admin
            '                If Trim(AccLock) = "ALL" Then
            '                    CekSch = "select * from tbltemplockacc where date(start_lock)='"
            '                    CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and f_locked isnull"
            '                Else
            '                    CekSch = "select * from tbltemplockacc where date(start_lock)='"
            '                    CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and account_lock='"
            '                    CekSch = CekSch + Trim(AccLock) + "' and f_locked isnull"
            '                End If
            '            End If
            '
                        
                        '@@diubah logikanya 14-12-2010
            ''
                        'Ini jika akan me-lock data semua
                        If (Trim(AccLock) = "ALL") Then
                            CekSch = "select * from mandiri.tbltemplockacc where f_locked isnull"
                        End If
                        'Ini jika akan me-lock data spv
                        If Left(Trim(AccLock), 3) = "SPV" Then
                            CekSch = "select * from mandiri.tbltemplockacc where f_locked isnull and (account_lock='ALL' or "
                            CekSch = CekSch + "account_lock in (select userid from mandiri.usertbl where spvcode='"
                            CekSch = CekSch + Trim(AccLock) + "') or account_lock='"
                            CekSch = CekSch + Trim(AccLock) + "')"
                        End If
                        If Trim(AccLock) <> "ALL" And Left(Trim(AccLock), 3) <> "SPV" Then
                            'Ini jika akan me-lock data per agent
                            CekSch = "select * from mandiri.tbltemplockacc where f_locked isnull and (account_lock='ALL' or "
                            CekSch = CekSch + "account_lock in (select spvcode from mandiri.usertbl where userid='"
                            CekSch = CekSch + Trim(AccLock) + "') or account_lock='"
                            CekSch = CekSch + Trim(AccLock) + "')"
                        End If
                        
                
                        
                        Set M_Objrs_CekSch = New ADODB.Recordset
                        M_Objrs_CekSch.CursorLocation = adUseClient
                        M_Objrs_CekSch.Open CekSch, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                        If M_Objrs_CekSch.RecordCount <> 0 Then
                            While Not M_Objrs_CekSch.EOF
                                
            '                    If WaktuAwal <= CDate(Format(M_Objrs_CekSch("end_lock"), "mm-dd-yyyy hh:mm")) _
            '                       Or WaktuAwal = CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
            '                    Then
                                    
                                 'Ubah @@ Logika 15-12-2010
                                 If (Format(WaktuAwal, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                                    And Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("end_lock"), "mm-dd-yyyy hh:mm"))) _
                                    Or _
                                    (Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                                    And Format(WaktuAkhir, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm"))) _
                                  Then
                                    'cek dulu deh, data lock dalam antrian di lock oleh siapa?
                                    'jika di lock sama spv, tl ga bisa release
                                    '@@ 22-11-10 ini di nonaktifkan dulu jadi team leader juga bisa release punya spv, ruquest mba wulan
                                    If UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                                        If Trim(M_Objrs_CekSch("lock_by")) = "WULAN" _
                                           Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SEPTIAN" _
                                           Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "JOKO" _
                                           Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "ADMIN" _
                                           Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "ADMINISTRATOR" _
                                           Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SUPERVISOR" Then
            
                                            pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                            pesan = pesan + " Karena agent ini telah di lock oleh SPV pada waktu yang bersamaan!" + Chr(13)
                                            pesan = pesan + " Silahkan hubungi Supervisor anda!"
            
                                            MsgBox pesan, vbOKOnly + vbExclamation, "Peringatan"
                                            Exit Sub
                                        End If
                                    End If
                                    '@@ 22-11-10 ini di nonaktifkan dulu jadi team leader juga bisa release punya spv, ruquest mba wulan
            
                                    
                                    'Jika yang lock pak septian maka wulan dan joko diblok
                                    If UCase(MDIForm1.TxtUsername.Text) = "WULAN" Or UCase(MDIForm1.TxtUsername.Text) = "JOKO" Or UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
                                        If UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SEPTIAN" Then
                                            pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                            pesan = pesan + " Karena agent ini telah di lock oleh Pak Septian pada waktu yang bersamaan!" + Chr(13)
                                            pesan = pesan + " Silahkan hubungi Pak Septian!"
                                            MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'Di lock dulu nih datanya biar ga bisa runing di agent
                                    LockStatus = "Update mandiri.tbltemplockacc set f_locked='5' where id='"
                                    LockStatus = LockStatus + Trim(M_Objrs_CekSch("id")) + "'"
                                    M_OBJCONN.Execute LockStatus
                                    
                                    'Kasih tau pesan ke user
                                    pesan = "PERHATIAN! :" + Chr(13)
                                    pesan = pesan + "Ada lock data dalam antrian yang sudah dibuat " + Chr(13)
                                    pesan = pesan + "dengan AGENT yang di lock = AGENT lock data yang akan dibuat, dan " + Chr(13)
                                    pesan = pesan + "WAKTU AWAL lock data yang dibuat lebih kecil dari  " + Chr(13)
                                    pesan = pesan + "WAKTU AKHIR lock data dalam antrian yang telah dibuat. " + Chr(13)
                                    pesan = pesan + "Hal ini dapat menimpa lock data yang telah dibuat. " + Chr(13)
                                    
                                    MsgBox pesan, vbOKOnly + vbCritical, "Peringatan"
                                    
                                    pesan = "Anda dapat menghapus lock data ini:" + Chr(13)
                                    pesan = pesan + "Start Lock: " + CStr(M_Objrs_CekSch("start_lock")) + Chr(13)
                                    pesan = pesan + "End Lock: " + CStr(M_Objrs_CekSch("end_lock")) + Chr(13)
                                    pesan = pesan + "Account yang di lock: " + M_Objrs_CekSch("account_lock") + Chr(13)
                                    pesan = pesan + "Di lock oleh: " + M_Objrs_CekSch("lock_by") + Chr(13)
                                    pesan = pesan + "Status yang di lock: " + M_Objrs_CekSch("status_lock") + Chr(13)
                                    pesan = pesan + Chr(13)
                                    pesan = pesan + "Jika anda ingin menyimpan schedule yang sedang dibuat, " + Chr(13)
                                    pesan = pesan + "anda harus menghapus data di atas!" + Chr(13)
                                    pesan = pesan + "Yakin lock data di atas akan dihapus?"
                                    
                                    a = MsgBox(pesan, vbYesNo + vbQuestion, "Konfirmasi")
                                    If a = vbYes Then
                                        Delsch = "delete from mandiri.tbltemplockacc where id='"
                                        Delsch = Delsch + Trim(M_Objrs_CekSch("id")) + "'"
                                        M_OBJCONN.Execute Delsch
                                        MsgBox "System akan kembali mengecek data! Jika ada data yang sama, system akan kembali mengkonfirmasi anda!", vbOKOnly + vbInformation, "Informasi"
                                    Else
                                        MsgBox "Lock data gagal ditambahkan!", vbOKOnly + vbExclamation, "Peringatan"
                                        'Buka lock data dalam antrian
                                        LockStatus = "update mandiri.tbltemplockacc set f_locked=null where id='"
                                        LockStatus = LockStatus + Trim(M_Objrs_CekSch("id")) + "'"
                                        M_OBJCONN.Execute LockStatus
                                        Set M_Objrs_CekSch = Nothing
                                        Exit Sub
                                    End If
                                    
                                End If
                                M_Objrs_CekSch.MoveNext
                            Wend
                        End If
                        Set M_Objrs_CekSch = Nothing
                        
                        
                        
                        'Cek apakah dalam tanggal waktu yang sama, ada lock data yang sedang RUNNING??
                        Dim CekSchCurrent As String
                        Dim M_Objrs_CekSchCurrent As ADODB.Recordset
                        
            '            'cek apakah dia sebagai teamleader or spv or admin
            '            If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
            '                 'Ini jika yang login TL
            '                 CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
            '                 CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' and  (lock_by='"
            '                 CekSchCurrent = CekSchCurrent + Trim(mdiform1.txtusername.text) + "' or lock_by in ('SEPTIAN','JOKO','WULAN','ADMIN','ADMINISTRATOR','SUPERVISOR')) "
            '                 CekSchCurrent = CekSchCurrent + " and account_lock in ('" + Trim(AccLock) + "','ALL')"
            '            Else
            '                'Ini jika yang login SPV atau Admin
            '                If Trim(AccLock) = "ALL" Then
            '                    CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
            '                    CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' "
            '                Else
            '                    CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
            '                    CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' and account_lock='"
            '                    CekSchCurrent = CekSchCurrent + Trim(AccLock) + "'"
            '                End If
            '            End If
                            
            ''            '@@diubah logikanya 14-12-2010
            ''            CekSchCurrent = "select * from tbltemplockacc_current where f_locked isnull and account_lock in ('ALL','"
            ''            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "','"
            ''            CekSchCurrent = CekSchCurrent + Trim(Replace(mdiform1.txtusername.text, "TL", "SPV")) + "')"
                            
                        'Ini untuk mengecek semua data
                        If (Trim(AccLock) = "ALL") Then
                            CekSchCurrent = "select * from tbltemplockacc_current where f_locked='1'"
                        End If
                        'Ini untuk mengecek data SPV
                        If Left(Trim(AccLock), 3) = "SPV" Then
                            CekSchCurrent = "select * from mandiri.tbltemplockacc_current where f_locked='1' and (account_lock='ALL' or "
                            CekSchCurrent = CekSchCurrent + "account_lock in (select userid from mandiri.usertbl where spvcode='"
                            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "') or account_lock='"
                            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "')"
                        End If
                        If Trim(AccLock) <> "ALL" And Left(Trim(AccLock), 3) <> "SPV" Then
                            'Ini untuk mengecek data per agent
                            CekSchCurrent = "select * from mandiri.tbltemplockacc_current where f_locked='1' and (account_lock='ALL' or "
                            CekSchCurrent = CekSchCurrent + "account_lock in (select spvcode from mandiri.usertbl where userid='"
                            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "') or account_lock='"
                            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "')"
                        End If
                        
                        
                        Set M_Objrs_CekSchCurrent = New ADODB.Recordset
                        M_Objrs_CekSchCurrent.CursorLocation = adUseClient
                        M_Objrs_CekSchCurrent.Open CekSchCurrent, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                        If M_Objrs_CekSchCurrent.RecordCount <> 0 Then
                            While Not M_Objrs_CekSchCurrent.EOF
                                
            '                    If WaktuAwal <= CDate(Format(M_Objrs_CekSchCurrent("end_lock"), "mm-dd-yyyy hh:mm")) _
            '                       Or WaktuAwal = CDate(Format(M_Objrs_CekSchCurrent("start_lock"), "mm-dd-yyyy hh:mm")) _
            '                    Then
                                    
                                 'Ubah @@ Logika 15-12-2010
                                 If (Format(WaktuAwal, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSchCurrent("start_lock"), "mm-dd-yyyy hh:mm")) _
                                    And Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSchCurrent("end_lock"), "mm-dd-yyyy hh:mm"))) _
                                    Or _
                                    (Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSchCurrent("start_lock"), "mm-dd-yyyy hh:mm")) _
                                    And Format(WaktuAkhir, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSchCurrent("start_lock"), "mm-dd-yyyy hh:mm"))) _
                                  Then
                                    
                                    'cek dulu deh, data lock dalam antrian di lock oleh siapa?
                                    'jika di lock sama spv, tl ga bisa release
                                    '@@22-11-10 INI DINONAKTIFKAN TERLEBIH DAHULU, REQUEST MBA WULAN
                                    If UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                                        If UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "WULAN" _
                                           Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SEPTIAN" _
                                           Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "JOKO" _
                                           Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "ADMIN" _
                                           Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "ADMINISTRATOR" _
                                           Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SUPERVISOR" Then
            
                                            pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                            pesan = pesan + " Karena waktu dari lock data dan agent yang di lock = waktu dari lock yang sedang berjalan!" + Chr(13)
                                            pesan = pesan + " Silahkan hubungi Supervisor anda!"
            
                                            MsgBox pesan, vbOKOnly + vbExclamation, "Peringatan"
                                            Exit Sub
                                        End If
                                    End If
                                    '@@22-11-10 INI DINONAKTIFKAN TERLEBIH DAHULU, REQUEST MBA WULAN
                                    
                                    
                                    'Jika yang lock pak septian maka wulan dan joko diblok
                                    If UCase(Trim(MDIForm1.TxtUsername.Text)) = "WULAN" Or UCase(Trim(MDIForm1.TxtUsername.Text)) = "JOKO" Or UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                                        If UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SEPTIAN" Then
                                            pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                            pesan = pesan + " Karena agent ini telah di lock oleh Pak Septian pada waktu yang bersamaan!" + Chr(13)
                                            pesan = pesan + " Silahkan hubungi Pak Septian!"
                                            MsgBox pesan, vbInformation + vbOKOnly, "Informasi"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    
                                    'Di lock dulu nih datanya biar ga bisa runing di agent
                                    'LockStatus = "Update tbltemplockacc_current set f_locked='6' where id='"
                                    'LockStatus = LockStatus + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                                    'M_OBJCONN.Execute LockStatus
                                    
                                    'Kasih tau pesan ke user
                                    pesan = "PERHATIAN! :" + Chr(13)
                                    pesan = pesan + "Ada lock data yang sedang RUNNING " + Chr(13)
                                    pesan = pesan + "dengan AGENT yang di lock = AGENT lock data yang akan dibuat, dan " + Chr(13)
                                    pesan = pesan + "WAKTU AWAL lock data yang dibuat lebih kecil dari  " + Chr(13)
                                    pesan = pesan + "WAKTU AKHIR lock data yang sedang RUNNING. " + Chr(13)
                                    pesan = pesan + "Hal ini dapat menimpa lock data yang telah dibuat. Harap RELEASE Terlebih dahulu lock data tersebut!" + Chr(13)
                                    
                                    MsgBox pesan, vbOKOnly + vbCritical, "Peringatan"
                                    
                                    pesan = "Anda dapat me-RELEASE lock data ini:" + Chr(13)
                                    pesan = pesan + "Start Lock: " + CStr(M_Objrs_CekSchCurrent("start_lock")) + Chr(13)
                                    pesan = pesan + "End Lock: " + CStr(M_Objrs_CekSchCurrent("end_lock")) + Chr(13)
                                    pesan = pesan + "Account yang di lock: " + M_Objrs_CekSchCurrent("account_lock") + Chr(13)
                                    pesan = pesan + "Di lock oleh: " + M_Objrs_CekSchCurrent("lock_by") + Chr(13)
                                    pesan = pesan + "Status yang di lock: " + M_Objrs_CekSchCurrent("status_lock") + Chr(13)
                                    pesan = pesan + Chr(13)
                                    pesan = pesan + "Jika anda ingin menyimpan schedule yang sedang dibuat, " + Chr(13)
                                    pesan = pesan + "anda harus me-RELEASE data di atas!" + Chr(13)
                                    pesan = pesan + "Yakin lock data di atas akan di RELEASE?"
                                    
                                    ' -- 22 APRIL 2014 HILANGKAN MESSAGE UNTUK SPV JADI BISA REPLACE LOC ACC --
                                    If MDIForm1.txtlevel.Text <> "Supervisor" Then
                                        a = MsgBox(pesan, vbYesNo + vbQuestion, "Konfirmasi")
                                    Else
                                        a = vbYes
                                    End If
                                    ' -------------------------------------------------------------------------
                                    
                                    If a = vbYes Then
                                    
                                         'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
                                        cmdsqlserver = "update mandiri.usertbl set dilockoleh='Clear by:" + Trim(MDIForm1.txtlevel.Text) + "',"
                                        cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
                                        cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
                                        'Buat ambil kondisi agent yang sedang di lock
                                        If Trim(M_Objrs_CekSchCurrent("account_lock")) = "ALL" Then
                                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
                                        ElseIf Left(Trim(M_Objrs_CekSchCurrent("account_lock")), 3) = "SPV" Then
                                            cmdsqlserver = cmdsqlserver + " where spvcode='"
                                            cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                                        Else
                                            cmdsqlserver = cmdsqlserver + " where userid='"
                                            cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                                        End If
                                        M_OBJCONN.Execute cmdsqlserver
                                    
                                    
                                        'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
                                        cmdsqlserver = "update mandiri.usertbl set f_pesanresetauto='1', f_pesanlockauto=null,f_idsessstart=null,f_idsessend=null "
                                        'Buat mengupdate pesan kondisi agent yang di lock
                                        If Trim(M_Objrs_CekSchCurrent("account_lock")) = "ALL" Then
                                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
                                        ElseIf Left(Trim(M_Objrs_CekSchCurrent("account_lock")), 3) = "SPV" Then
                                            cmdsqlserver = cmdsqlserver + " where spvcode='"
                                            cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                                        Else
                                            cmdsqlserver = cmdsqlserver + " where userid='"
                                            cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                                        End If
                                        M_OBJCONN.Execute cmdsqlserver
                                        
                                        'Clossing Session
                                        Dim UpdateDtCloseSession As String
                                        UpdateDtCloseSession = "update mandiri.tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
                                        UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
                                        UpdateDtCloseSession = UpdateDtCloseSession + " from mandiri.tblperformpersessionlock inner join mandiri.mgm "
                                        UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
                                        UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
                                        UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
                                        UpdateDtCloseSession = UpdateDtCloseSession + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                                        M_OBJCONN.Execute UpdateDtCloseSession
                                        'Akhir dari closing session
                                        
                                        'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
                                        cmdsqlserver = "insert into mandiri.tbltemplockacc_log select * from mandiri.tbltemplockacc_current where "
                                        cmdsqlserver = cmdsqlserver + " id='"
                                        cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                                        M_OBJCONN.Execute cmdsqlserver
                                    
                                    
                                        Delsch = "delete from mandiri.tbltemplockacc_current where id='"
                                        Delsch = Delsch + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                                        M_OBJCONN.Execute Delsch
                                        
                                        If MDIForm1.txtlevel.Text <> "Supervisor" Then
                                            MsgBox "System akan kembali mengecek data! Jika ada data yang sama, system akan kembali mengkonfirmasi anda!", vbOKOnly + vbInformation, "Informasi"
                                        End If
                                    Else
                                        MsgBox "Lock data gagal ditambahkan!", vbOKOnly + vbExclamation, "Peringatan"
                                        'Buka lock data dalam antrian
                                        'LockStatus = "update tbltemplockacc_current set f_locked=null where id='"
                                        'LockStatus = LockStatus + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                                        'M_OBJCONN.Execute LockStatus
                                        Set M_Objrs_CekSchCurrent = Nothing
                                        Exit Sub
                                    End If
                                    
                                End If
                                M_Objrs_CekSchCurrent.MoveNext
                            Wend
                        End If
                        Set M_Objrs_CekSchCurrent = Nothing
                    
            
                        CmdTimer = "insert into mandiri.tbltemplockacc (date_lock,start_lock,end_lock,"
                        CmdTimer = CmdTimer + "account_lock,lock_by,status_lock,script_lock) values ('"
                        CmdTimer = CmdTimer + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
                        CmdTimer = CmdTimer + Format(StartDate.Value, "yyyy-mm-dd") + " " + Format(StartTime.Value, "hh:mm:00") + "','"
                        CmdTimer = CmdTimer + Format(EndDate.Value, "yyyy-mm-dd") + " " + Format(EndTime.Value, "hh:mm:00") + "','"
                        CmdTimer = CmdTimer + Trim(AccLock) + "','"
                        CmdTimer = CmdTimer + Trim(MDIForm1.TxtUsername.Text) + "','"
                        CmdTimer = CmdTimer + Trim(StatusLocked) + Trim(HLSMARKUP) + "-" + "','"
                        CmdTimer = CmdTimer + Trim(StringBlokTimer) + "')"
                        M_OBJCONN.Execute CmdTimer
                    End If
                Next z
                
                
                MsgBox "Lock data berhasil ditambahkan dalam schedulle!", vbOKOnly + vbInformation, "Informasi"
                  
                StsVl = ""
                StsPR = ""
                StsOS = ""
                StsON = ""
                StsSK = ""
                StsOP = ""
                StsPTP = ""
                StsBP = ""
                StsPOP = ""
                StsSP = ""
                StsUC = ""
                StsRP = ""
                StsWO_Date = ""
                StsWO_2009 = ""
                StsWO_2008 = ""
                StsWO_2007 = ""
                StsWO_2006 = ""
                StsWO_2005 = ""
                StsWO_2004 = ""
                StsWO_2003 = ""
                StsWO_2002 = ""
                StsWO_2001 = ""
                StsWO_2000 = ""
                StsWO_1999 = ""
                STRSQL = ""
                '@@ 18-11-10 perbaikan blok data entry
                StsNewEntry = ""
                StsReguler = ""
                StsCurrent = ""
                StsSwap = ""
                
                Unload Me
                '@@ 061110 - akhir blok dengan timer -
     
    Case 1
         If SSOption1(2).Value = True Then
                   
                If CHKLUNASPENDING.Value = vbChecked Or CHKACCOUNT.Value = vbChecked Then
                        
                       Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                
                If Combo1(0).Text = "" Then
                        M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE usertype='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
                Else
                        M_objrs.Open "SELECT * FROM mandiri.USERTBL WHERE SPVCODE='" + Combo1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
                End If
                
                While Not M_objrs.EOF
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispvbuattl=NULL WHERE SPVCODE='" + M_objrs("SPVCODE") + "'"
                    M_OBJCONN.Execute STRSQL
                    M_objrs.MoveNext
                Wend
                Set M_objrs = Nothing
                MsgBox "Data telah direlease"
               Exit Sub
               End If
                
                    If Combo1(0).Text = "" Then
                    MsgBox "CLIK DULU COMBO SPV", vbInformation + vbOKOnly, "PESAN"
                    Exit Sub
                   End If
                        STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='Clear by:"
                        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                        STRSQL = STRSQL + " lockdarispv=NULL,lock_ptp_payment=NULL,lock_entry_lpd=NULL,fromaccount=NULL,lockmarkup=NULL,lockdarispvbuattl=NULL WHERE SPVCODE='" + Combo1(0).Text + "'"
                    
         Else
                If SSOption1(1).Value = True Then
                    If Combo1(0).Text = "" Then
                        MsgBox "CLIK DULU COMBO NYA", vbInformation + vbOKOnly, "PESAN"
                    Exit Sub
                    End If
                    STRSQL = "UPDATE mandiri.usertbl SET dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,lock_ptp_payment=NULL,lockmarkup=NULL ,fromaccount=NULL,lockdarispvbuattl=NULL WHERE userid='" + Combo1(0).Text + "'"
                Else
                        STRSQL = "UPDATE mandiri.usertbl SET lockdarispv=NULL,lock_entry_lpd=NULL,lock_ptp_payment=NULL,lockmarkup=NULL,dilockoleh='Clear by:"
                        STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "-" + Format(Now, "yyyy-mm-dd") + "'"
                End If
                
        End If
        If STRSQL <> "" Then
            M_OBJCONN.Execute STRSQL
            MsgBox "Reset Done.....!"
            End If
            
    Case 2
            Unload Me
    
    Case 3
            STRSQL1 = "UPDATE  mandiri.tblshut SET nshut=1 "
            M_OBJCONN.Execute STRSQL1
            
    Case 4
            Call preview_blok
    End Select

End Sub

Private Sub preview_blok()
    Dim m_preview_selected As ADODB.Recordset
    Dim sqlstr As String
    
    Set m_preview_selected = New ADODB.Recordset
    m_preview_selected.ActiveConnection = M_OBJCONN
    m_preview_selected.CursorLocation = adUseClient
    m_preview_selected.CursorType = adOpenDynamic
    
    If Check1(0).Value = 1 Or _
        Check1(1).Value = 1 Or _
        Check1(2).Value = 1 Or _
        Check1(3).Value = 1 Or _
        Check1(4).Value = 1 Or _
        Check1(5).Value = 1 Or _
        Check1(6).Value = 1 Or _
        Check1(7).Value = 1 Or _
        Check1(9).Value = 1 Or _
        Check1(10).Value = 1 Or _
        Check1(11).Value = 1 Then
            Call ceksts
    Else
        Text1.Text = ""
        ListView3.ListItems.CLEAR
        MsgBox "Pilih Status Data yang akan di Lock terlebih dahulu!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    sqlstr = "SELECT custid FROM mandiri.mgm WHERE custid IS NOT NULL "
    
    sqlstr = sqlstr & IIf(getblock <> "", " AND " & Replace(getblock, "@", "'"), "")
    sqlstr = sqlstr & IIf(getBlokPTPNoPayment <> "", " AND " & Replace(getBlokPTPNoPayment, "@", "'"), "")
    sqlstr = sqlstr & IIf(GetBlockEntry <> "", " AND " & Replace(GetBlockEntry, "@", "'"), "")
    
    m_preview_selected.Open Replace(sqlstr, "''", "''")
    
    ListView3.ListItems.CLEAR
    If m_preview_selected.RecordCount > 0 Then
        While Not m_preview_selected.EOF
            ListView3.ListItems.ADD , , cnull(m_preview_selected!CustId)
            m_preview_selected.MoveNext
        Wend
    End If

    Text1.Text = m_preview_selected.RecordCount
    Set m_preview_selected = Nothing
End Sub

Sub ceksts()
StsVl = ""
If Check1(0) = vbChecked Then
    StsVl = "VL-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsVl
End If

StsOP = ""
If Check1(1).Value Then
    StsOP = "OS-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsOP
End If

StsPTP = ""
If Check1(2).Value Then
    StsPTP = "PTP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPTP + "-"
End If

StsBP = ""
If Check1(3).Value Then
    StsBP = "BP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsBP
End If

StsPOP = ""
If Check1(4).Value Then
    StsPOP = "POP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPOP
End If
 
StsSP = ""
If Check1(5).Value Then
    StsSP = "SP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSP
End If

StsRP = ""
If Check1(7).Value Then
    StsRP = "RP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsRP
End If

Stsblank = ""
If Check1(6).Value Then
    Stsblank = "anto"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "BlankData -"
End If

StsPR = ""
If Check1(9).Value Then
    StsPR = "PR-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPR
End If

StsON = ""
If Check1(10).Value Then
    StsON = "ON-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsON
End If

StsSK = ""
If Check1(11).Value Then
    StsSK = "SK-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSK
End If

If chknewentry.Value Then
    StatusLocked = StatusLocked + "NewEntry -"
    
    StsNewEntry = " ("
    StsNewEntry = StsNewEntry + "date_part(''month'',entry_date) between (date_part(''month'',now())- 2) "
    StsNewEntry = StsNewEntry + " and date_part(''month'',now())-1 "
    StsNewEntry = StsNewEntry + " and date_part(''year'',entry_date)=date_part(''year'',now()) "
    StsNewEntry = StsNewEntry + " )"

End If

If chkreguler.Value Then
    StatusLocked = StatusLocked + "Reguler -"
    
    StsReguler = " ("
    StsReguler = StsReguler + "date_part(''month'',pay_dt_update) between (date_part(''month'',now())- 2) "
    StsReguler = StsReguler + " and date_part(''month'',now())-1 "
    StsReguler = StsReguler + " and date_part(''year'',pay_dt_update)=date_part(''year'',now())"
    StsReguler = StsReguler + " )"
    
End If


If chkswap.Value Then
    StatusLocked = StatusLocked + "Swap -"
    
    StsSwap = " ("
    StsSwap = StsSwap + "( (date_part(''month'',pay_dt_update) < (date_part(''month'',now())- 2) "
    StsSwap = StsSwap + " and date_part(''year'',pay_dt_update) <= date_part(''year'',now())) "
    StsSwap = StsSwap + " or pay_dt_update isnull ) "
    StsSwap = StsSwap + " and "
    StsSwap = StsSwap + " date_part(''month'',entry_date) < (date_part(''month'',now())-2) "
    StsSwap = StsSwap + " and date_part(''year'',entry_date) <= date_part(''year'',now()) "
    StsSwap = StsSwap + " )"
End If


If chkcurrent.Value Then
    StatusLocked = StatusLocked + "Current -"
    
    StsCurrent = " ("
    StsCurrent = StsCurrent + " date_part(''month'',tglsource)=date_part(''month'',now()) "
    StsCurrent = StsCurrent + " and date_part(''year'',tglsource)=date_part(''year'',now()) "
    StsCurrent = StsCurrent + " )"
End If


'@@ 15 Agustus 2011, Request Gaby buat lock payment yang dari PIL dipindah ke card
If CeKBlokPayment.Value = vbChecked Then
    'Blok 1 bulan lalu
    If OptPayment1Bln.Value = True Then
        StatusLocked = StatusLocked + "BlokPayment1BlnLalu-"
        StsPayment1Bln = "= (date_part(''month'',now()) - 1)"
    End If
    
    'Blok 2 bulan lalu
    If OptPayment2Bln.Value = True Then
        StatusLocked = StatusLocked + "BlokPayment2BlnLalu-"
        StsPayment2Bln = "= (date_part(''month'',now())- 2)"
    End If
    
    'Blok 3 bulan lalu
    If OptPayment3Bln.Value = True Then
        StatusLocked = StatusLocked + "BlokPayment3BlnLalu-"
        StsPayment2Bln = "= (date_part(''month'',now())- 3)"
    End If
    
    'Blok > 3 bulan lalu
    If OptLbh3Bln.Value = True Then
        StatusLocked = StatusLocked + "BlokPayment<3BlnLalu-"
        StsPayment2Bln = "< (date_part(''month'',now())- 3)"
    End If
End If


'--- @@ 18-11-10 blok dulu deh buat diperbaiki skripnya u/ yg blok entry---------------
'BlokEntry = ""
'bCheckNewentry = False
'bCheckReguler = False
'bCheckSwap = False
'bCheckCurrent = False
'
'If chknewentry.Value = vbChecked Then
'    bCheckNewentry = True
'    '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "NewEntry -"
'End If
'
'
'If chkreguler.Value = vbChecked Then
'   bCheckReguler = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "Reguler -"
'End If
'
'If chkswap.Value = vbChecked Then
'   bCheckSwap = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "SWAP -"
'End If
'
'
'If chkcurrent.Value = vbChecked Then
'   bCheckCurrent = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "Current -"
'End If
'
'
'
'
'
'
'If bCheckSwap = True And bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + " date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckNewentry = True And bCheckSwap = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckReguler = True And bCheckSwap = True And bCheckCurrent = True Then
'   BlokEntry = " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now) or "
'   BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'   BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'   BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'   BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'   BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'   BlokEntry = BlokEntry + " )"
'   Exit Sub
'End If
'
'
'
'
'If bCheckNewentry = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'    Exit Sub
'End If
'
'
'If bCheckReguler = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'    Exit Sub
'End If
'
'
'If bCheckSwap = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'    Exit Sub
'End If
'
'If bCheckCurrent = True Then
'   BlokEntry = " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    Exit Sub
'End If
'--- @@ 18-11-10 blok dulu deh buat diperbaiki skripnya u/ yg blok entry---------------


'@@ 140710 Tambahan buat blok entry

'If OptNewEntry.Value = True Then
'    BlokEntry = " datediff(''month'',entry_date,now) "
'    BlokEntry = BlokEntry + CmbNewEntry.Text
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbReguler.Text
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbSwap.Text
'End If

'@@ 150710 Ubah blok entry
'If OptNewEntry.Value = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'End If

End Sub
Public Function getblock() As String


                    STRINGBLOK = ""
                    
                    If StsVl <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                            STRINGBLOK = " substring(F_cek_new,1,3) in (@" + StsVl + "@"
                        Else
                            STRINGBLOK = STRINGBLOK + ",@" + StsVl + "@"
                        End If
                    End If
                    
                    If StsPR <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPR + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPR + "@"
                        End If
                    End If
                    
                    If StsPTP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPTP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPTP + "@"
                        End If
                    End If
                    
                    If StsPOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPOP + "@"
                        End If
                    End If
                    
                    If StsBP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsBP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsBP + "@"
                        End If
                    End If
                    
                    If StsSP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)  in (@" + StsSP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSP + "@"
                        End If
                    End If
                    
                    If StsRP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsRP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsRP + "@"
                        End If
                    End If
                    
                    If StsSK <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsSK + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSK + "@"
                        End If
                    End If
                    
                     If StsON <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsON + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsON + "@"
                        End If
                    End If
                    
                     If StsOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsOP + "@"
                        End If
                    End If
                    
                    
                     If Stsblank <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_NEW,1,3)   in (@@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@@"
                        End If
                    End If
                    
                    
                    
                
                    If Len(STRINGBLOK) > 0 Then
                            STRINGBLOK = STRINGBLOK + ")"
                    End If
                    getblock = STRINGBLOK
End Function
Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Dim M_objrs As ADODB.Recordset
Dim cmdsqluser As String

Select Case Index
Case 0
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
Case 1
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        
        '@@221010
'        Dim M_DATA As New CLS_FRMSEARCH
'        Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
'            While Not M_OBJRS.EOF
'                Combo1(0).AddItem M_OBJRS("USERID")
'                Combo1(1).AddItem M_OBJRS("AGENT")
'                M_OBJRS.MoveNext
'            Wend
'        Set M_OBJRS = Nothing

       '@@ 11-11-10 tambahan kode jika tl yang menggunakan
        If Left(Trim(MDIForm1.TxtUsername.Text), 2) = "TL" Then
            cmdsqluser = "select * from mandiri.usertbl where usertype='1' and spvcode='"
            cmdsqluser = cmdsqluser + Replace(Trim(MDIForm1.TxtUsername.Text), "TL", "SPV") + "'"
            cmdsqluser = cmdsqluser + " order by userid asc"
        Else
           cmdsqluser = "select * from mandiri.usertbl where usertype='1' order by userid asc"
        End If
       
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsqluser, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("userid")
            Combo1(1).AddItem M_objrs("agent")
            M_objrs.MoveNext
        Wend
        
        'SSOption1(0).Value = True
        spv = False
Case 2
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            M_objrs.Open "select * from mandiri.SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
            M_objrs.Open "select * from mandiri.SPVTBL", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "select * from mandiri.SPVTBL where team='" + MDIForm1.TxtUsername.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        End If
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("SPVCODE")
                Combo1(1).AddItem M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
'        SSOption1(0).Value = True
'        SSOption1(1).Value = True
        
End Select

End Sub
Public Sub getMarkup()
Dim list As ListItem
Dim RSNEW As New ADODB.Recordset
Set rs = New ADODB.Recordset
RSNEW.CursorLocation = adUseClient
RSNEW.Open "select distinct(exclude) from mgm WHERE (exclude <>'')", M_OBJCONN, adOpenDynamic, adLockOptimistic
ListView1.ListItems.CLEAR
While Not RSNEW.EOF
Set list = ListView1.ListItems.ADD(, , IIf(IsNull(RSNEW!exclude), "", RSNEW!exclude))
    RSNEW.MoveNext
Wend

End Sub

Public Function GETSELECTMARKUP() As String
Dim j As Integer
Dim TMPSELECTMARKUP As String
GETSELECTMARKUP = ""
For j = 1 To ListView2.ListItems.Count
        If j = 1 Then
            TMPSELECTMARKUP = TMPSELECTMARKUP + Chr(39) + ListView2.ListItems(j).Text + Chr(39)
        Else
            TMPSELECTMARKUP = TMPSELECTMARKUP + "," + Chr(39) + ListView2.ListItems(j).Text + Chr(39)
        End If
    
        
Next j
GETSELECTMARKUP = TMPSELECTMARKUP
End Function

'@@ 18-11-10 ini perbaikan script blok entry data
Public Function GetBlockEntry() As String
    Dim StringBlokEntry As String


                    StringBlokEntry = ""
                    
                    If StsNewEntry <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                            StringBlokEntry = StsNewEntry
                        Else
                            StringBlokEntry = StringBlokEntry + " or " + StsNewEntry
                        End If
                    End If
                    
                    If StsReguler <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsReguler
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsReguler
                        End If
                    End If
                    
                    If StsSwap <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsSwap
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsSwap
                        End If
                    End If
                    
                    If StsCurrent <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsCurrent
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsCurrent
                        End If
                    End If
                    
                    'Ini buat ngasih kurung buka dan kurung tutup blok entry
                    If StringBlokEntry <> "" Then
                        StringBlokEntry = "( " + StringBlokEntry
                        StringBlokEntry = StringBlokEntry + " )"
                    End If
                    
                    GetBlockEntry = StringBlokEntry
End Function

Private Sub CekValidNullLock()
     If (IsNull(StsVl) Or StsVl = "") _
        And (IsNull(StsOS) Or StsOS = "") _
        And (IsNull(StsON) Or StsON = "") _
        And (IsNull(StsSK) Or StsSK = "") _
        And (IsNull(StsPR) Or StsPR = "") _
        And (IsNull(StsPTP) Or StsPTP = "") _
        And (IsNull(StsBP) Or StsBP = "") _
        And (IsNull(StsPOP) Or StsPOP = "") _
        And (IsNull(StsSP) Or StsSP = "") _
        And (IsNull(StsRP) Or StsRP = "") _
        And (IsNull(StsOP) Or StsOP = "") _
        And (IsNull(Stsblank) Or Stsblank = "") _
        And (IsNull(Stsuncontact) Or Stsuncontact = "") _
        And (IsNull(StsNewEntry) Or StsNewEntry = "") _
        And (IsNull(StsReguler) Or StsReguler = "") _
        And (IsNull(StsSwap) Or StsSwap = "") _
        And (IsNull(StsCurrent) Or StsCurrent = "") _
        And (ListView2.ListItems.Count = 0) _
        And (StatusLocked = "") Then
        
        MsgBox "Anda belum memilih, status data yang akan di lock!", vbOKOnly + vbInformation, "Informasi"
        CekValidLock = False
        Exit Sub
     End If
'     If StatusLocked = "" Then
'        MsgBox "Anda belum memilih, status data yang akan di lock!", vbOKOnly + vbInformation, "Informasi"
'        CekValidLock = False
'     End If
     CekValidLock = True
End Sub



Private Sub Timer_stopwatch_Timer()
    
    'Tambah dengan satu untuk total sepersepuluh detik.
    'Kita mengeset interval Timer menjadi 10, jadi
    'setiap sepersepuluh detik prosedur ini akan
    'dieksekusi
    TotalTenthDetik = TotalTenthDetik + 1
    'Jika TotalTenthSeconds = 10,
    'set kembali menjadi 0.
    TenthDetik = TotalTenthDetik Mod 10
    '10 kali sepersepuluh detik sama dengan 1 detik.
    'int - akan mengembalikan bilangan integer (bulat)
    'dari pecahan 'Contoh: Int(0.9) = 0 menghasilkan 0
    TotalDetik = Int(TotalTenthDetik / 10)
    'Jika variabel Seconds = 60, set kembali menjadi 0
    Detik = TotalDetik Mod 60
    If Len(Detik) = 1 Then
       Detik = "0" & Detik  'Agar selalu dalam dua
                            'digit
    End If
    Menit = Int(TotalDetik / 60) Mod 60
    If Len(Menit) = 1 Then
       Menit = "0" & Menit    'Agar selalu dalam dua
                          'digit
    End If
    JAM = Int(TotalDetik / 3600)
    If JAM < 9 Then
       Jam1 = "0" & JAM       'Agar selalu dalam dua'digit
    End If
    'Tampilkan hasilnya di Lblwaktu (update terus Lblwaktu)
    LblWaktu.Caption = Jam1 & ":" & Menit & ":" & Detik & ":" & TenthDetik & ""
End Sub

'@@ 01/03/2011 ,, User yang di lock bisa dipilih
Private Sub headerUser()
    LvUser.ColumnHeaders.ADD , , "User", 1000
    LvUser.ColumnHeaders.ADD , , "Name", 2000
    LvUser.ColumnHeaders.ADD , , "SPV Code", 1000
End Sub

Private Sub IsiUser()
    Dim ListItem As ListItem
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = "select * from mandiri.usertbl where usertype='1' and spvcode='"
        CMDSQL = CMDSQL + Trim(Replace(MDIForm1.TxtUsername.Text, "TL", "SPV")) + "'"
        CMDSQL = CMDSQL + " order by spvcode,userid asc"
    End If
    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or _
        UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or _
        UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "MANAGER" Then
        
        CMDSQL = "Select * from mandiri.usertbl where usertype='1'"
        CMDSQL = CMDSQL + " order by spvcode,userid asc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvUser.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvUser.ListItems.ADD(, , Trim(M_objrs("userid")))
                ListItem.SubItems(1) = Trim(M_objrs("agent"))
                ListItem.SubItems(2) = Trim(cnull(M_objrs("spvcode")))
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
    
End Sub

'@@ 01-03-2011 Clear data sebelum dilock
Private Sub ClearData()
    Dim i As Integer
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Data user tidak ada!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For i = 1 To LvUser.ListItems.Count
        If LvUser.ListItems(i).Checked = True Then
            STRSQL = "UPDATE mandiri.usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
            STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
            STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL Where userid='" + Trim(LvUser.ListItems(i).Text) + "'"
            M_OBJCONN.Execute (STRSQL)
        End If
    Next i
End Sub

Private Sub IsiSpv()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    CMDSQL = "select * from mandiri.spvtbl order by spvcode asc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        CmbSpv.CLEAR
        If M_objrs.RecordCount > 0 Then
            While Not M_objrs.EOF
                CmbSpv.AddItem Trim(M_objrs("spvcode"))
                M_objrs.MoveNext
            Wend
        End If
    Set M_objrs = Nothing
End Sub


'@@ 15 Agustus 2011, Request Gaby blok payment
Public Function GetBlokPayment() As String
    Dim StringBlokPayment As String
    
    If OptPayment1Bln.Value = False And _
       OptPayment2Bln.Value = False And _
       OptPayment3Bln.Value = False And _
       OptLbh3Bln.Value = False Then
       
       StringBlokPayment = ""
       GetBlokPayment = StringBlokPayment
       Exit Function
    End If
       
    
    StringBlokPayment = " custid in (select custid from mandiri.tbllunas where  "
    StringBlokPayment = StringBlokPayment + "date_part(''month'',paydate) "
    
    If StsPayment1Bln <> "" Then
        StringBlokPayment = StringBlokPayment + StsPayment1Bln
    End If
                    
    If StsPayment2Bln <> "" Then
        StringBlokPayment = StringBlokPayment + StsPayment2Bln
    End If
                    
    If StsPayment3Bln <> "" Then
        StringBlokPayment = StringBlokPayment + StsPayment3Bln
    End If
    
    If StsPaymentLbh3Bln <> "" Then
        StringBlokPayment = StringBlokPayment + StsPaymentLbh3Bln
    End If
    
    StringBlokPayment = StringBlokPayment + " and date_part(''year'',paydate)=date_part(''year'',now()) ) "
    'StringBlokPayment = StringBlokPayment + " id in (select max(id) as id from tbllunas group by custid,agent) ) "
    
    GetBlokPayment = StringBlokPayment
End Function

Function getBlokPTPNoPayment() As String
    If chk_ptp_payment.Value = vbChecked Then
        getBlokPTPNoPayment = " custid in (SELECT a.custid FROM mandiri.reportptp a LEFT JOIN " & _
                        "(SELECT custid,max(paydate) as tgl_akhir,payment as Total FROM mandiri.tbllunas WHERE date_part(''year'',paydate)=date_part(''year'',now()) AND date_part(''month'',paydate)=date_part(''month'',now()) GROUP BY" & _
                        " custid,payment) b ON a.custid=b.custid WHERE date_part(''month'',a.inputdate)=date_part(''month'',now()) AND date_part(''year'',a.inputdate)=date_part(''year'',now()) AND a.promisedate<=now()) "
    Else
        getBlokPTPNoPayment = ""
    End If
End Function
