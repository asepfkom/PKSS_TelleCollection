VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_deskcoll_performance 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DeskColl Performance Insentif"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13470
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid msfx 
      Height          =   5655
      Left            =   240
      TabIndex        =   25
      Top             =   1440
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   9975
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FCFCFC&
      Height          =   2055
      Left            =   240
      TabIndex        =   5
      Top             =   7440
      Width           =   12975
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "Payment No PTP"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   11
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   9
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   8
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0"
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   7
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   10
         Left            =   12060
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   6
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   5
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   4
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   3
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   2
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   600
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   240
         Width           =   800
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1515
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Export to Excel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1150
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Selisih"
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
         Index           =   11
         Left            =   3720
         TabIndex        =   41
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "All Payment"
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
         Index           =   9
         Left            =   3720
         TabIndex        =   39
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BP Amount"
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
         Index           =   8
         Left            =   360
         TabIndex        =   37
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept Amount"
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
         Index           =   7
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   12960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Data"
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
         Index           =   10
         Left            =   10800
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept + Broken"
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
         Index           =   6
         Left            =   7080
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kept"
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
         Index           =   5
         Left            =   7080
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RPC 2"
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
         Index           =   4
         Left            =   5040
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RPC 1"
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
         Index           =   3
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PTP"
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
         Index           =   2
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dialer Hours"
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
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Hours"
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
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCFCFC&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin MSComDlg.CommonDialog CD 
         Left            =   6840
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5520
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   8040
         TabIndex        =   28
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cb_sort 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4440
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker tgl_laporan 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM-yyyy"
         Format          =   96141315
         CurrentDate     =   41610
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Upload Absensi"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
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
         Left            =   3840
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sort Column By"
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
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan dan Tahun"
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   6240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ket : RPC 1 : PTP,KP,BP,ON,PR     RPC 2 : CH, SPOUSE "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   43
      Top             =   7200
      Width           =   7455
   End
End
Attribute VB_Name = "Form_deskcoll_performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_calc As ADODB.Recordset
Private rs_temp As ADODB.Recordset

Private tgl_lap As Date
Private sql_str As String

Private sPaid_hours As Long
Private sDialer_hours As Long
Private sPTP As Long
Private sRPC1 As Long
Private sRPC2 As Long
Private sKept As Long
Private sKept_old As Long
Private sKeptBroken As Long
Private sKeptBrokenOld As Long
Private sKeptAmount As Double
Private sPTP_Reg As Double
Private sPymnGlobal As Double

Private tPaid_hours As Long
Private tDialer_hours As Long
Private tPTP As Long
Private tRPC1 As Long
Private tRPC2 As Long
Private tKept As Long
Private tKept_old As Long
Private tKeptBroken As Long
Private tKeptBrokenOld As Long
Private tKeptAmount As Double
Private tPTP_Reg As Double
Private tPymnGlobal As Double

Private sqlfilter As String
Private m_SortColumn As Integer
Private m_SortOrder As Integer

Private sql_tahun       As String
Private sql_bulan       As String

Private connICENTRA4 As ADODB.Connection
Private connICENTRA5 As ADODB.Connection

Private Sub cb_sort_Click()
    SortByColumn cb_sort.ListIndex
End Sub

Private Sub Combo1_Click()
    ' ---- OPSI AGENT ----
    If Combo1.Text <> "" And Combo1.Text <> "ALL" Then
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' AND team='" & Combo1.Text & "' ORDER BY userid"
        Combo2.CLEAR
        Combo3.CLEAR
        Do Until rs_temp.EOF
            Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
            Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
            rs_temp.MoveNext
        Loop
        ' -------------------
    Else
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' ORDER BY userid"
        Combo2.CLEAR
        Combo3.CLEAR
        Do Until rs_temp.EOF
            Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
            Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
            rs_temp.MoveNext
        Loop
        ' -------------------
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    Combo3.ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo3_Click()
    Combo2.ListIndex = Combo3.ListIndex
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

'Private Sub Command1_Click()
'    Dim lstItem         As ListItem
'    Dim xx              As Integer
'    Dim warna_belang    As Integer
'
'    sqlFilter = ""
'    Command1.Enabled = False
'
'    If Combo1.Text <> "" And Combo1.Text <> "ALL" Then
'        sqlFilter = " AND a.team='" & Combo1.Text & "'"
'    End If
'
'    If Combo2.Text <> "" Then
'        sqlFilter = sqlFilter & " AND a.userid='" & Combo2.Text & "'"
'    End If
'
'    tgl_lap = Format(tgl_laporan.Value, "yyyy-mm-dd")
'    sql_str = ""
'    If rs_calc.state = 1 Then rs_calc.Close
'    ' OLD 16 JAN 2013 ==============
''    sql_str = "SELECT a.userid,c.agent as nama,c.team as TL, floor(a.jml_jam) as Paid_hours,floor(d.jml_dialer) as Dialer_hours,b.jml_ptp as PTP,j.ptp_regular, f.sts1 as RPC1, g.sts2 as RPC2 ,coalesce(e.jml_kept,0) as Kept, coalesce(k.jml_cust) as Kept_old, coalesce(e.jml_kept,0) + coalesce(h.broken,0) as ""Kept+Broken"", coalesce(k.jml_cust,0) + coalesce(l.jml_broken_old,0) as ""Kept+Broken_OLD"", coalesce(e.keptamount,0) as kept_amount,coalesce(jmlpayment_global,0) as payment_global " & _
''                " FROM "
''    ' PAID HOURS
''    sql_str = sql_str + " (SELECT y.userid,sum(x.hours) as jml_jam FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY userid,nopeg ) a LEFT JOIN "
''    ' JUMLAH PTP
''    sql_str = sql_str + " (SELECT agent,count(custid) as jml_PTP FROM (SELECT distinct agent,custid FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY agent,custid) GROUP BY agent) b ON a.userid=b.agent LEFT JOIN "
''    ' JUMLAH PTP REGULAR
''    sql_str = sql_str + " (SELECT xx.agent,count(xx.custid) as PTP_Regular FROM (SELECT distinct agent,custid FROM tblnegoptp_log WHERE date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY agent,custid) xx,mgm yy WHERE xx.custid=yy.custid GROUP BY xx.agent) j ON a.userid=j.agent, "
''    ' DIALER HOURS
''    sql_str = sql_str + " usertbl c,(SELECT userid,sum(hours) as jml_dialer FROM tblabsen_aplikasi WHERE date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY userid ) d LEFT JOIN "
''    ' KEPT PROMISE
''    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as jml_kept,sum(y.payment) as keptamount FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY custid,agent,promisepay) x, " & _
''                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y WHERE x.custid=y.custid AND (y.paydate between x.Tgl_janji AND x.Tgl_janji+5) AND y.payment>=x.promisepay GROUP BY x.agent) e ON d.userid=e.agent LEFT JOIN "
''    ' BROKEN PROMISE(BP)
''    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as broken FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY custid,agent,promisepay) x LEFT JOIN " & _
''                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y ON x.custid=y.custid WHERE ((y.paydate between x.Tgl_janji AND x.Tgl_janji+5) OR (date(x.Tgl_janji)+5=date(now())) ) AND (y.payment<x.promisepay OR y.payment is null) GROUP BY x.agent) h ON d.userid=h.agent, "
''    ' BROKEN OLD
''    sql_str = sql_str + " (SELECT distinct agent,count(custid) as jml_broken_old FROM tblnegoptp_log WHERE custid not in ( SELECT distinct custid FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) AND promisedate between '" & Format(tgl_lap, "yyyy-mm-01") & "' and '" & Format(Now, "yyyy-mm-dd") & "' GROUP BY agent) l ,"
''    ' RPC 1
''    sql_str = sql_str + " (SELECT agent,count(id) as sts1 FROM mgm WHERE substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP by agent) f, "
''    ' RPC 2
''    sql_str = sql_str + " (SELECT agent,count(id) as sts2 FROM mgm WHERE trim(lower(statuscall)) in ('ch','spouse') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP by agent) g, "
''    ' Payment Global
''    sql_str = sql_str + " (SELECT agent,count(custid) as jml_cust,sum(payment) as jmlpayment_global FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY agent) k"
''    sql_str = sql_str + " WHERE a.userid=c.userid AND a.userid=d.userid AND a.userid=f.agent AND a.userid=g.agent AND a.userid=k.agent AND a.userid=l.agent " & sqlFilter & " ORDER BY a.userid;"
'
'    sql_str = "SELECT a.userid,a.agent as nama,a.team as TL, floor(jml_jam) as Paid_hours,floor(jml_dialer) as Dialer_hours,jml_ptp as PTP,ptp_regular, sts1 as RPC1, sts2 as RPC2 ,coalesce(jml_kept,0) as Kept, coalesce(jml_cust,0) as Kept_old, coalesce(jml_kept,0) + coalesce(broken,0) as ""Kept+Broken"", coalesce(jml_cust,0) + coalesce(jml_broken_old,0) as ""Kept+Broken_OLD"", coalesce(keptamount,0) as kept_amount,coalesce(jmlpayment_global,0) as payment_global FROM usertbl a LEFT JOIN "
'    ' PAID HOURS
'    sql_str = sql_str + " (SELECT dd.*,ee.jml_cust,jmlpayment_global FROM (SELECT bb.*,cc.sts2 FROM (SELECT z.*,aa.sts1 FROM (SELECT x.*,y.jml_broken_old FROM (SELECT v.*,w.broken FROM (SELECT t.*,u.jml_kept,u.keptamount FROM (SELECT r.*,s.PTP_Regular FROM (SELECT userid,jml_jam,jml_dialer,jml_ptp FROM (SELECT o.*,p.jml_dialer FROM (SELECT m.userid, n.jml_jam FROM usertbl m LEFT JOIN(SELECT y.userid,sum(x.hours) as jml_jam FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY userid,nopeg ) n on m.userid=n.userid) o LEFT JOIN  "
'    ' DIALER HOURS
'    sql_str = sql_str + " (SELECT userid,sum(hours) as jml_dialer FROM tblabsen_aplikasi WHERE date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY userid ) p ON o.userid=p.userid WHERE o.userid like 'D%') p LEFT JOIN "
'    ' JUMLAH PTP
'    sql_str = sql_str + " (SELECT agent,count(custid) as jml_PTP FROM (SELECT distinct agent,custid FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) GROUP BY agent) q ON p.userid=q.agent ) r LEFT JOIN "
'    ' JUMLAH PTP REGULAR
'    sql_str = sql_str + " (SELECT xx.agent,count(xx.custid) as PTP_Regular FROM (SELECT distinct agent,custid FROM tblnegoptp_log WHERE date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY agent,custid) xx GROUP BY xx.agent) s ON r.userid=s.agent) t LEFT JOIN "
'    ' KEPT PROMISE
'    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as jml_kept,sum(y.payment) as keptamount FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "' ) AND (date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) GROUP BY custid,agent,promisepay) x, " & _
'                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y WHERE x.custid=y.custid AND y.payment>=x.promisepay AND (y.paydate>=x.Tgl_janji OR y.paydate between x.Tgl_janji-3 AND x.Tgl_janji) GROUP BY x.agent) u ON t.userid=u.agent) v LEFT JOIN  "
'    ' BROKEN PROMISE(BP)
'    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as broken FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE (date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) AND (date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) GROUP BY custid,agent,promisepay) x LEFT JOIN " & _
'                        " (SELECT custid,agent,paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "') y ON x.custid=y.custid WHERE x.agent=y.agent AND (y.payment<x.promisepay OR y.payment is null) AND (y.paydate>=x.Tgl_janji OR y.paydate is null) GROUP BY x.agent) w ON v.userid=w.agent) x LEFT JOIN "
'    ' BROKEN OLD
'    'sql_str = sql_str + " (SELECT distinct agent,count(custid) as jml_broken_old FROM tblnegoptp_log WHERE custid not in ( SELECT distinct custid FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) AND promisedate between '" & Format(tgl_lap, "yyyy-mm-01") & "' and '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lap, "yyyy-mm-01"))), "yyyy-mm-dd") & "' GROUP BY agent) y ON x.userid=y.agent) z LEFT JOIN "
'    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as jml_broken_old FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE (date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) GROUP BY custid,agent,promisepay) x LEFT JOIN " & _
'                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y ON x.custid=y.custid WHERE (y.payment<x.promisepay OR y.payment is null) AND (y.paydate>=x.Tgl_janji OR y.paydate is null) GROUP BY x.agent) y ON x.userid=y.agent) z LEFT JOIN "
'
'    ' RPC 1
'    sql_str = sql_str + " (SELECT agent,count(id) as sts1 FROM mgm WHERE substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP by agent) aa ON z.userid=aa.agent) bb LEFT JOIN "
'    ' RPC 2
'    sql_str = sql_str + " (SELECT agent,count(id) as sts2 FROM mgm WHERE trim(lower(statuscall)) in ('ch','spouse') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP by agent) cc ON bb.userid=cc.agent) dd LEFT JOIN "
'    ' Payment Global
'    'sql_str = sql_str + " (SELECT agent,count(custid) as jml_cust,sum(payment) as jmlpayment_global FROM tbllunas WHERE custid in (SELECT distinct custid FROM tblnegoptp_log WHERE (promisedate between '" & Format(tgl_lap, "yyyy-mm-01") & "' and '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lap, "yyyy-mm-01"))), "yyyy-mm-dd") & "')) AND date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY agent) ee ON dd.userid=ee.agent) b ON a.userid=b.userid "
'    sql_str = sql_str + " (SELECT x.agent,count(x.custid) as jml_cust,sum(y.payment) as jmlpayment_global FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY custid,agent,promisepay) x, " & _
'                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y WHERE x.custid=y.custid AND y.payment>=x.promisepay AND (y.paydate>=x.Tgl_janji OR y.paydate between x.Tgl_janji-3 AND x.Tgl_janji) GROUP BY x.agent) ee ON dd.userid=ee.agent)b on a.userid=b.userid  "
'
'    sql_str = sql_str + " WHERE a.userid like 'D%' " & sqlFilter & " ORDER BY a.userid;"
'
'    rs_calc.Open sql_str
'
'    sPaid_hours = 0
'    sDialer_hours = 0
'    sPTP = 0
'    sRPC1 = 0
'    sRPC2 = 0
'    sKept = 0
'    sKept_old = 0
'    sKeptBroken = 0
'    sKeptBrokenOld = 0
'    sKeptAmount = 0
'    sPTP_Reg = 0
'    sPymnGlobal = 0
'
'    msfx.Rows = 1
'    warna_belang = 0
'    'ListView1.ListItems.CLEAR
'    If rs_calc.RecordCount > 0 Then
'        'Set DataGrid1.DATASOURCE = rs_calc
'        Command3.Enabled = True
'        ProgressBar1.Max = rs_calc.RecordCount
'        While Not rs_calc.EOF
'            DoEvents
'            ProgressBar1.Value = rs_calc.Bookmark
'            tPaid_hours = Val(cnull(rs_calc!paid_hours))
'            tDialer_hours = Val(cnull(rs_calc!Dialer_hours))
'            tPTP = Val(cnull(rs_calc!ptp))
'            tPTP_Reg = Val(cnull(rs_calc!ptp_regular))
'            tRPC1 = Val(cnull(rs_calc!rpc1))
'            tRPC2 = Val(cnull(rs_calc!rpc2))
'            tKept = Val(cnull(rs_calc!kept))
'            tKept_old = Val(cnull(rs_calc!kept_old))
'            tKeptBroken = Val(cnull(rs_calc("Kept+Broken")))
'            tKeptBrokenOld = Val(cnull(rs_calc("Kept+Broken_old")))
'            tKeptAmount = Val(cnull(rs_calc("Kept_amount")))
'            tPymnGlobal = Val(cnull(rs_calc("payment_global")))
'
'            With msfx
'                .Rows = .Rows + 1
'                xx = rs_calc.Bookmark
'                .TextMatrix(xx, 1) = xx
'                .TextMatrix(xx, 2) = cnull(rs_calc!Userid)
'                .TextMatrix(xx, 3) = cnull(rs_calc!Nama)
'                .TextMatrix(xx, 4) = cnull(rs_calc!TL)
'                .TextMatrix(xx, 5) = tPaid_hours
'                .TextMatrix(xx, 6) = tDialer_hours
'                .TextMatrix(xx, 7) = tPTP
'                .TextMatrix(xx, 8) = tPTP_Reg
'                .TextMatrix(xx, 9) = tRPC1
'                .TextMatrix(xx, 10) = tRPC2
'                .TextMatrix(xx, 11) = tKept
'                .TextMatrix(xx, 12) = tKept_old
'                .TextMatrix(xx, 13) = tKeptBroken
'                .TextMatrix(xx, 14) = tKeptBrokenOld
'                .TextMatrix(xx, 15) = Format(tKeptAmount, "#,###,###")
'                .TextMatrix(xx, 16) = Format(tPymnGlobal, "#,###,###")
'
'                ' Dialer Op / Paid Hours
'                If tDialer_hours > 0 And tPaid_hours > 0 Then
'                    .TextMatrix(xx, 17) = Round(tDialer_hours / tPaid_hours, 2)
'                End If
'
'                ' RPC / Dialer
'                If tRPC2 > 0 And tDialer_hours > 0 Then
'                    .TextMatrix(xx, 18) = Round(tRPC2 / tDialer_hours, 2)
'                End If
'
'                ' RPC / Paid Hours
'                If tRPC2 > 0 And tPaid_hours > 0 Then
'                    .TextMatrix(xx, 19) = Round(tRPC2 / tPaid_hours, 2)
'                End If
'
'                ' KeptBroken / RPC
'                If tKeptBroken > 0 And tRPC2 > 0 Then
'                    .TextMatrix(xx, 20) = Round(tKeptBroken / tRPC2, 2)
'                End If
'
'                ' PTP / RPC
'                If tPTP > 0 And tRPC2 > 0 Then
'                    .TextMatrix(xx, 21) = Round(tPTP / tRPC2, 2)
'                End If
'
'                ' KEPT / KEPT BROKEN
'                If tKept > 0 And tKeptBroken > 0 Then
'                    .TextMatrix(xx, 22) = Round(tKept / tKeptBroken, 2)
'                End If
'
'                ' Average Payment Size
'                If tKeptAmount > 0 And tKept > 0 Then
'                    .TextMatrix(xx, 23) = Format(Round(tKeptAmount / tKept, 0), "#,###,###")
'                End If
'
'                ' CEV
'                .TextMatrix(xx, 24) = Format(Round(Val(.TextMatrix(xx, 20)) * Val(.TextMatrix(xx, 22)) * Val(Format(.TextMatrix(xx, 23), "#")), 0), "#,###,###")
'
'                ' EVPH
'                .TextMatrix(xx, 25) = Format(Round(Val(.TextMatrix(xx, 19)) * Val(Format(.TextMatrix(xx, 24), "#")), 0), "#,###,###")
'
'
'
'                ' ############################## OLD CALCULATION #############################
'
'                ' KeptBroken / RPC OLD
'                If tKeptBroken > 0 And tRPC2 > 0 Then
'                    .TextMatrix(xx, 26) = Round(tKeptBrokenOld / tRPC2, 2)
'                End If
'
'                ' PTP / RPC OLD
'                If tPTP > 0 And tRPC2 > 0 Then
'                    .TextMatrix(xx, 27) = Round(tPTP / tRPC2, 2)
'                End If
'
'                ' KEPT / KEPT BROKEN OLD
'                If tKept_old > 0 And tKeptBrokenOld > 0 Then
'                    .TextMatrix(xx, 28) = Round(tKept_old / tKeptBrokenOld, 2)
'                End If
'
'                ' Average Payment Size
'                If tPymnGlobal > 0 And tKept_old > 0 Then
'                    .TextMatrix(xx, 29) = Format(Round(tPymnGlobal / tKept_old, 0), "#,###,###")
'                End If
'
'                ' CEV
'                .TextMatrix(xx, 30) = Format(Round(Val(.TextMatrix(xx, 26)) * Val(.TextMatrix(xx, 28)) * Val(Format(.TextMatrix(xx, 29), "#")), 0), "#,###,###")
'                ' EVPH
'                .TextMatrix(xx, 31) = Format(Round(Val(.TextMatrix(xx, 19)) * Val(Format(.TextMatrix(xx, 30), "#")), 0), "#,###,###")
'
'                ' ############################## OLD CALCULATION ############################
'
'
'                If warna_belang = 1 Then
'                    For i = 1 To msfx.Cols - 1
'                        .Col = i
'                        .Row = xx
'                        .CellBackColor = &HEFEFEF
'                    Next i
'                    warna_belang = 0
'                Else
'                    warna_belang = 1
'                End If
'            End With
'
'            sPaid_hours = sPaid_hours + Val(cnull(rs_calc!paid_hours))
'            sDialer_hours = sDialer_hours + Val(cnull(rs_calc!Dialer_hours))
'            sPTP = sPTP + Val(cnull(rs_calc!ptp))
'            sPTP_Reg = sPTP_Reg + Val(cnull(rs_calc!ptp_regular))
'            sRPC1 = sRPC1 + Val(cnull(rs_calc!rpc1))
'            sRPC2 = sRPC2 + Val(cnull(rs_calc!rpc2))
'            sKept = sKept + Val(cnull(rs_calc!kept))
'            sKept_old = sKept_old + Val(cnull(rs_calc!kept_old))
'            sKeptBroken = sKeptBroken + Val(cnull(rs_calc("Kept+Broken")))
'            sKeptBrokenOld = sKeptBrokenOld + Val(cnull(rs_calc("Kept+Broken_old")))
'            sKeptAmount = sKeptAmount + Val(cnull(rs_calc("Kept_amount")))
'            sPymnGlobal = sPymnGlobal + tPymnGlobal
'
'            rs_calc.MoveNext
'        Wend
'        rs_calc.MoveFirst
'    Else
'        Command3.Enabled = False
'        MsgBox "Hasil data tidak ada!!", vbOKOnly + vbInformation, "INFO"
'    End If
'
'    ' ======= TOTAL DATA ======
'    Text1(0).Text = Format(sPaid_hours, "#,###,###")
'    Text1(1).Text = Format(sDialer_hours, "#,###,###")
'    Text1(2).Text = Format(sPTP, "#,###,###")
'    Text1(3).Text = Format(sRPC1, "#,###,###")
'    Text1(4).Text = Format(sRPC2, "#,###,###")
'    Text1(5).Text = Format(sKept, "#,###,###")
'    Text1(6).Text = Format(sKeptBroken, "#,###,###")
'    Text1(7).Text = Format(sKeptAmount, "#,###,###")
'    Text1(8).Text = Format(sPTP_Reg, "#,###,###")
'    Text1(9).Text = Format(sPymnGlobal, "#,###,###")
'    Text1(10).Text = Format(rs_calc.RecordCount, "#,###,###")
'    Text1(11).Text = Format(sKept_old, "#,###,###")
'    Text1(12).Text = Format(sKeptBrokenOld, "#,###,###")
'    ' =========================
'    Command1.Enabled = True
'
'End Sub

Private Sub Command1_Click()
    Dim lstItem         As listItem
    Dim xx              As Integer
    Dim warna_belang    As Integer
    
    Dim temp_payment    As Double
    Dim temp_paydate    As String
    
    Dim f_kept          As String
    
    Dim dial_paidh As Double
    Dim rpc_dialh As Double
    Dim rpc_paidh As Double
    Dim keptbroken_rpc As Double
    Dim ptp_rpc As Double
    Dim kept_keptbroken As Double
    Dim avg_paymentsize As Double
    Dim cev As Double
    Dim evph As Double

    sqlfilter = ""
    Command1.Enabled = False
    
    tgl_lap = Format(tgl_laporan.Value, "yyyy-mm-dd")
    sql_str = ""

    sql_tahun = Format(tgl_lap, "yyyy")
    sql_bulan = Format(tgl_lap, "mm")
    
    M_OBJCONN.Execute "DELETE FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & ""
    
    ' INSERT PTP
    M_OBJCONN.Execute "INSERT INTO tblreport_komisi(custid,promisedate,agent,tglinput,promisepay,tgl_report) " & _
                    " SELECT custid,promisedate,agent,tglinput,promisepay,to_date('" & Format(tgl_lap, "yyyy-mm-01") & "','YYYY-MM-DD') " & _
                    " FROM tblnegoptp_log  WHERE date_part('month',tglinput)=" & sql_bulan & " AND date_part('year',tglinput)=" & sql_tahun & " AND date_part('month',promisedate)=" & sql_bulan & " AND date_part('year',promisedate)=" & sql_tahun & ";"
    
    ' UPDATE PTP->PAYMENT untuk agent terakhir PTP
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT id,custid,promisedate,agent,promisepay FROM tblreport_komisi"
    If rs_temp.RecordCount > 0 Then
        ProgressBar1.Max = rs_temp.RecordCount
        Do Until rs_temp.EOF
            DoEvents
            ProgressBar1.Value = rs_temp.Bookmark
         
            sql_str = "SELECT sum(a.payment) as total_bayar,max(a.paydate) as tgl_akhir FROM (SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & sql_bulan & " AND date_part('year',paydate)=" & sql_tahun & " ) a " & _
                        " WHERE a.custid='" & cnull(rs_temp!CustId) & "'"
            
            
            temp_paydate = "null"
            temp_payment = 0
            
            If rs_calc.state = 1 Then rs_calc.Close
            rs_calc.Open sql_str
            If rs_calc.RecordCount > 0 Then
               temp_paydate = IIf(IsNull(rs_calc!tgl_akhir), "Null", "'" & Format(rs_calc!tgl_akhir, "yyyy-mm-dd") & "'")
               temp_payment = IIf(IsNull(rs_calc!total_bayar), 0, rs_calc!total_bayar)
            End If
            
            If temp_payment >= cnull(rs_temp!PromisePay) Then
                f_kept = "KEPT"
            Else
                f_kept = "BP"
            End If
            
            ' UPDATE untuk agent yang terakhir ---------------
            M_OBJCONN.Execute "UPDATE tblreport_komisi SET payment=" & temp_payment & ",paydate=" & temp_paydate & ",f_kept='" & f_kept & "' FROM (SELECT custid,max(promisedate) as tgl_akhirPTP FROM tblreport_komisi WHERE date(tgl_report)='" & Format(tgl_lap, "yyyy-mm-01") & "' AND custid='" & cnull(rs_temp!CustId) & "' GROUP BY custid) a WHERE tblreport_komisi.custid=a.custid AND promisedate=a.tgl_akhirPTP "
            
            rs_temp.MoveNext
        Loop
    End If
    
    M_OBJCONN.Execute "DELETE FROM result_komisi ;"
    
    sql_str = "INSERT INTO result_komisi(userid,name,tl,paid_hours,dialer_hours,ptp,rpc1,rpc2,kept,keptbroken,keptamount) SELECT userid,agent as nama,team as TL, floor(jml_jam) as Paid_hours,floor(jml_dialer) as Dialer_hours,jml_ptp as PTP, sts1 as RPC1, sts2 as RPC2 ,coalesce(jml_kept,0) as Kept, coalesce(jml_kept,0) + coalesce(broken,0) as ""Kept+Broken"", coalesce(keptamount,0) as kept_amount FROM " & _
                "(SELECT z.*,aa.sts2 FROM (SELECT x.*,y.sts1 FROM (SELECT v.*,w.broken FROM (SELECT t.*,u.jml_kept,u.keptamount FROM (SELECT r.*,s.jml_ptp FROM (SELECT o.*,p.jml_dialer FROM (SELECT a.userid,a.agent,a.team,b.jml_jam FROM (SELECT * FROM usertbl WHERE usertype='1' AND userid like 'D%') a LEFT JOIN "
    ' PAID HOURS
    sql_str = sql_str + " (SELECT y.userid,sum(x.hours) as jml_jam FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=" & sql_bulan & " AND date_part('year',tanggal)=" & sql_tahun & " GROUP BY userid,nopeg) b ON a.userid=b.userid) o LEFT JOIN  "
    ' DIALER HOURS
    sql_str = sql_str + " (SELECT userid,sum(hours) as jml_dialer FROM tblabsen_aplikasi WHERE date_part('month',tanggal)=" & sql_bulan & " AND date_part('year',tanggal)=" & sql_tahun & " GROUP BY userid ) p ON o.userid=p.userid) r LEFT JOIN "
    ' JUMLAH PTP
    sql_str = sql_str + " (SELECT agent,count(agent) as jml_PTP FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " GROUP BY agent) s ON r.userid=s.agent ) t LEFT JOIN "
    ' KEPT
    sql_str = sql_str + " (SELECT agent,count(agent) as jml_kept,sum(payment) as keptamount FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " AND f_kept='KEPT' GROUP BY agent) u ON t.userid=u.agent ) v LEFT JOIN "
    ' BP
    sql_str = sql_str + " (SELECT agent,count(agent) as broken FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " AND f_kept='BP' GROUP BY agent) w ON v.userid=w.agent ) x LEFT JOIN "
    ' RPC 1
    ' sql_str = sql_str + " (SELECT agent,count(id) as sts1 FROM mgm_hst WHERE substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') AND agent like 'D%' AND date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " GROUP by agent) y ON x.userid=y.agent) z LEFT JOIN "
    sql_str = sql_str + " (SELECT agent,count(a.agent) as sts1 FROM (SELECT custid,agent,tgl,ststelpwith FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-')) a,(SELECT custid, max(tgl) as tgl_akhir FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') GROUP BY custid ) b WHERE a.custid=b.custid AND a.tgl=b.tgl_akhir GROUP BY agent) y ON x.userid=y.agent) z LEFT JOIN "
    ' RPC 2
    sql_str = sql_str + " (SELECT agent,count(a.agent) as sts2 FROM (SELECT custid,agent,tgl,ststelpwith FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND ststelpwith in ('CH','SPOUSE')) a,(SELECT custid, max(tgl) as tgl_akhir FROM mgm_hst WHERE date_part('month',tgl)=" & sql_bulan & " AND date_part('year',tgl)=" & sql_tahun & " AND ststelpwith in ('CH','SPOUSE') GROUP BY custid ) b WHERE a.custid=b.custid AND a.tgl=b.tgl_akhir GROUP BY agent) aa ON z.userid=aa.agent) ORDER BY userid "

    M_OBJCONN.Execute sql_str
    
    If Combo1.Text <> "" And Combo1.Text <> "ALL" Then
        sqlfilter = " AND tl='" & Combo1.Text & "'"
    End If

    If Combo2.Text <> "" Then
        sqlfilter = sqlfilter & " AND userid='" & Combo2.Text & "'"
    End If
    
    sPaid_hours = 0
    sDialer_hours = 0
    sPTP = 0
    sRPC1 = 0
    sRPC2 = 0
    sKept = 0
    sKeptBroken = 0
    sKeptAmount = 0
    
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT * FROM result_komisi WHERE userid is not null " & sqlfilter & " ORDER BY userid"
    If rs_temp.RecordCount > 0 Then
        ProgressBar1.Max = rs_temp.RecordCount
        warna_belang = 0
        msfx.Rows = 1
        msfx.Rows = rs_temp.RecordCount + 1
        
        Do Until rs_temp.EOF
            DoEvents
            ProgressBar1.Value = rs_temp.Bookmark
            tPaid_hours = Val(cnull(rs_temp!paid_hours))
            tDialer_hours = Val(cnull(rs_temp!Dialer_hours))
            tPTP = Val(cnull(rs_temp!ptp))
            tRPC1 = Val(cnull(rs_temp!rpc1))
            tRPC2 = Val(cnull(rs_temp!rpc2))
            tKept = Val(cnull(rs_temp!kept))
            tKeptBroken = Val(cnull(rs_temp("Keptbroken")))
            tKeptAmount = Val(cnull(rs_temp("Keptamount")))

            With msfx
                xx = rs_temp.Bookmark
                .TextMatrix(xx, 1) = xx
                .TextMatrix(xx, 2) = cnull(rs_temp!Userid)
                .TextMatrix(xx, 3) = cnull(rs_temp!Name)
                .TextMatrix(xx, 4) = cnull(rs_temp!TL)
                .TextMatrix(xx, 5) = tPaid_hours
                .TextMatrix(xx, 6) = tDialer_hours
                .TextMatrix(xx, 7) = tPTP
                .TextMatrix(xx, 8) = tRPC1
                .TextMatrix(xx, 9) = tRPC2
                .TextMatrix(xx, 10) = tKept
                .TextMatrix(xx, 11) = tKeptBroken
                .TextMatrix(xx, 12) = Format(tKeptAmount, "#,###,###")
                
                ' Dialer Op / Paid Hours
                If tDialer_hours > 0 And tPaid_hours > 0 Then
                    dial_paidh = Round(tDialer_hours / tPaid_hours, 2)
                    .TextMatrix(xx, 13) = dial_paidh
                End If

                ' RPC / Dialer
                If tRPC2 > 0 And tDialer_hours > 0 Then
                    rpc_dialh = Round(tRPC2 / tDialer_hours, 2)
                    .TextMatrix(xx, 14) = rpc_dialh
                End If

                ' RPC / Paid Hours
                If tRPC2 > 0 And tPaid_hours > 0 Then
                    rpc_paidh = Round(tRPC2 / tPaid_hours, 2)
                    .TextMatrix(xx, 15) = rpc_paidh
                End If

                ' KeptBroken / RPC
                If tKeptBroken > 0 And tRPC2 > 0 Then
                    keptbroken_rpc = Round(tKeptBroken / tRPC2, 2)
                    .TextMatrix(xx, 16) = keptbroken_rpc
                End If

                ' PTP / RPC
                If tPTP > 0 And tRPC2 > 0 Then
                    ptp_rpc = Round(tPTP / tRPC2, 2)
                    .TextMatrix(xx, 17) = ptp_rpc
                End If

                ' KEPT / KEPT BROKEN
                If tKept > 0 And tKeptBroken > 0 Then
                    kept_keptbroken = Round(tKept / tKeptBroken, 2)
                    .TextMatrix(xx, 18) = kept_keptbroken
                End If

                ' Average Payment Size
                If tKeptAmount > 0 And tKept > 0 Then
                    avg_paymentsize = Format(Round(tKeptAmount / tKept, 0), "#,###,###")
                    .TextMatrix(xx, 19) = avg_paymentsize
                End If

                ' CEV
                cev = Round(Val(.TextMatrix(xx, 16)) * Val(.TextMatrix(xx, 18)) * Val(Format(.TextMatrix(xx, 19), "#")), 0)
                .TextMatrix(xx, 20) = Format(cev, "#,###,###")

                ' EVPH
                evph = Round(Val(.TextMatrix(xx, 15)) * Val(Format(.TextMatrix(xx, 20), "#")), 0)
                .TextMatrix(xx, 21) = Format(evph, "#,###,###")

                sPaid_hours = sPaid_hours + Val(cnull(rs_temp!paid_hours))
                sDialer_hours = sDialer_hours + Val(cnull(rs_temp!Dialer_hours))
                sPTP = sPTP + Val(cnull(rs_temp!ptp))
                sRPC1 = sRPC1 + Val(cnull(rs_temp!rpc1))
                sRPC2 = sRPC2 + Val(cnull(rs_temp!rpc2))
                sKept = sKept + Val(cnull(rs_temp!kept))
                sKeptBroken = sKeptBroken + Val(cnull(rs_temp("KeptBroken")))
                sKeptAmount = sKeptAmount + Val(cnull(rs_temp("Keptamount")))
                
                If warna_belang = 1 Then
                    For i = 1 To msfx.Cols - 1
                        .Col = i
                        .Row = xx
                        .CellBackColor = &HEFEFEF
                    Next i
                    warna_belang = 0
                Else
                    warna_belang = 1
                End If
                
                M_OBJCONN.Execute "UPDATE result_komisi SET dial_paidh=" & dial_paidh & ",rpc_dialh=" & rpc_dialh & ",rpc_paidh=" & rpc_paidh & "," & _
                                "keptbroken_rpc=" & keptbroken_rpc & ",ptp_rpc=" & ptp_rpc & ",kept_keptbroken=" & kept_keptbroken & ",avg_paymentsize=" & avg_paymentsize & ",cev=" & cev & ",epvh=" & evph & " WHERE " & _
                                "userid='" & rs_temp!Userid & "'"
                
                rs_temp.MoveNext
            End With
        Loop
    Else
        MsgBox "Data tidak ditemukan!!", vbOKOnly + vbInformation, "INFO"
    End If
    
    ' ======= TOTAL DATA ======
    Text1(0).Text = Format(sPaid_hours, "#,###,###")
    Text1(1).Text = Format(sDialer_hours, "#,###,###")
    Text1(2).Text = Format(sPTP, "#,###,###")
    Text1(3).Text = Format(sRPC1, "#,###,###")
    Text1(4).Text = Format(sRPC2, "#,###,###")
    Text1(5).Text = Format(sKept, "#,###,###")
    Text1(6).Text = Format(sKeptBroken, "#,###,###")
    Text1(7).Text = Format(sKeptAmount, "#,###,###")
    Text1(10).Text = Format(rs_temp.RecordCount, "#,###,###")
    ' =========================
    
    If rs_temp.state = 1 Then rs_temp.Close
    ' -- BP AMOUNT
    rs_temp.Open " SELECT sum(payment) as total_bp FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & " AND f_kept='BP'"
    If rs_temp.RecordCount > 0 Then
        DoEvents
        Text1(8).Text = Format(rs_temp!total_bp, "#,###,###")
    End If
    
    If rs_temp.state = 1 Then rs_temp.Close
    ' -- ALL PAYMENT
    rs_temp.Open "SELECT sum(payment) as total_all FROM tbllunas WHERE  date_part('month',paydate)=" & sql_bulan & " AND date_part('year',paydate)=" & sql_tahun
    
    If rs_temp.RecordCount > 0 Then
        DoEvents
        Text1(9).Text = Format(rs_temp!total_all, "#,###,###")
    End If
    
    If rs_temp.state = 1 Then rs_temp.Close
    ' -- SELISIH
    rs_temp.Open "SELECT sum(payment) as total_selisih FROM tbllunas WHERE  date_part('month',paydate)=" & sql_bulan & " AND date_part('year',paydate)=" & sql_tahun & " and custid not in( " & _
                "SELECT distinct custid FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & ");"
    
    If rs_temp.RecordCount > 0 Then
        DoEvents
        Text1(11).Text = Format(rs_temp!total_selisih, "#,###,###")
    End If
    
    Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    Form_upload_absensi.Show 1
End Sub

Private Sub Command3_Click()
    CD.Filter = "Excel Files (*.xls)|*.xls"
    CD.ShowSave
    If CD.FileName <> "" Then
        If rs_calc.state = 1 Then rs_calc.Close
        rs_calc.Open "SELECT * FROM result_komisi ORDER BY agent;"
        
        If rs_calc.RecordCount > 0 Then
            ConvertToExcel rs_calc, CD.FileName
        Else
            MsgBox "Tidak ada data yang didownload!!", vbOKOnly + vbInformation, "INFO"
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    Dim spath As String
    
    CD.Filter = "Excel Files (*.xls)|*.xls"
    CD.ShowSave
    spath = CD.FileName
    If spath <> "" Then
        If rs_temp.state = 1 Then rs_temp.Close
        ' -- SELISIH
        rs_temp.Open "SELECT custid,agent,paydate,payment FROM tbllunas WHERE  date_part('month',paydate)=" & sql_bulan & " AND date_part('year',paydate)=" & sql_tahun & " and custid not in( " & _
                    "SELECT distinct custid FROM tblreport_komisi WHERE date_part('month',tgl_report)=" & sql_bulan & " AND date_part('year',tgl_report)=" & sql_tahun & ");"
        
        If rs_temp.RecordCount > 0 Then
            Call ConvertToExcel(rs_temp, spath)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call koneksi
    
    ' ---- OPSI TL ----
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT distinct team as team_TL FROM usertbl WHERE team is not null AND lower(team) not in ('reserved','septian','wulan','admin')"
    Combo1.CLEAR
    Combo1.AddItem "ALL"
    Do Until rs_temp.EOF
        Combo1.AddItem IIf(IsNull(rs_temp!team_TL), "", rs_temp!team_TL)
        rs_temp.MoveNext
    Loop
    ' ------------------
    
    ' ---- OPSI AGENT ----
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT userid,agent FROM usertbl WHERE userid like 'D%' ORDER BY userid"
    Combo2.CLEAR
    Combo3.CLEAR
    Do Until rs_temp.EOF
        Combo2.AddItem IIf(IsNull(rs_temp!Userid), "", rs_temp!Userid)
        Combo3.AddItem IIf(IsNull(rs_temp!agent), "", rs_temp!agent)
        rs_temp.MoveNext
    Loop
    ' -------------------
    
    msfx.Cols = 22
    With msfx
        .TextMatrix(0, 1) = "No"
        .TextMatrix(0, 2) = "User ID"
        .TextMatrix(0, 3) = "Name"
        .TextMatrix(0, 4) = "TL"
        .TextMatrix(0, 5) = "Paid Hours"
        .TextMatrix(0, 6) = "Dialer Hours"
        .TextMatrix(0, 7) = "PTP"
        .TextMatrix(0, 8) = "RPC 1"
        .TextMatrix(0, 9) = "RPC 2"
        .TextMatrix(0, 10) = "Kept"
        .TextMatrix(0, 11) = "Kept+Broken"
        .TextMatrix(0, 12) = "Kept Amount"
        .TextMatrix(0, 13) = "Dialer Op/Paid Hrs"
        .TextMatrix(0, 14) = "RPC/Dialer Op Hrs"
        .TextMatrix(0, 15) = "RPC/Paid Hrs"
        .TextMatrix(0, 16) = "(Kept+Broken)/RPC"
        .TextMatrix(0, 17) = "PTP/RPC"
        .TextMatrix(0, 18) = "Kept#/(Kept+Broken)"
        .TextMatrix(0, 19) = "Average Payment Size"
        .TextMatrix(0, 20) = "CEV"
        .TextMatrix(0, 21) = "EVPH"
    End With
    
'    ListView1.ColumnHeaders(1).Width = 720
'    ListView1.ColumnHeaders(2).Width = 720
'    ListView1.ColumnHeaders(3).Width = 2445
'    ListView1.ColumnHeaders(4).Width = 720
'    ListView1.ColumnHeaders(5).Width = 950
'    ListView1.ColumnHeaders(6).Width = 950
'    ListView1.ColumnHeaders(7).Width = 950
'    ListView1.ColumnHeaders(8).Width = 950
'    ListView1.ColumnHeaders(9).Width = 950
'    ListView1.ColumnHeaders(10).Width = 950
'    ListView1.ColumnHeaders(11).Width = 1000
'    ListView1.ColumnHeaders(12).Width = 1750
    For i = 0 To msfx.Cols - 1
        cb_sort.AddItem msfx.TextMatrix(0, i)
    Next i
End Sub

Private Sub koneksi()
    Set rs_calc = New ADODB.Recordset
    rs_calc.CursorLocation = adUseClient
    rs_calc.CursorType = adOpenDynamic
    rs_calc.LockType = adLockOptimistic
    rs_calc.ActiveConnection = M_OBJCONN
    
    Set rs_temp = New ADODB.Recordset
    rs_temp.CursorLocation = adUseClient
    rs_temp.CursorType = adOpenDynamic
    rs_temp.LockType = adLockOptimistic
    rs_temp.ActiveConnection = M_OBJCONN
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_temp = Nothing
    Set rs_calc = Nothing
End Sub

Private Sub ListView1_Click()
'msgbox listview1.SelectedItem.SubItems(listview1.SelectedItem
End Sub

Private Sub msfx_DblClick()
    Dim strUser_Selected As String
    Dim sql_selected As String
    Dim strCaption As String
    
    strUser_Selected = msfx.TextMatrix(msfx.Row, 2)
    sql_selected = ""
    Select Case msfx.Col
    Case 5 ' PAID HOURS
        strCaption = "PAID HOURS"
        sql_selected = "SELECT userid,agent,masuk,keluar,tanggal,hours as jml_jam FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND y.userid='" & strUser_Selected & "'"
    Case 6 ' DIALER HOURS
        strCaption = "DIALER HOURS"
        sql_selected = "SELECT * FROM tblabsen_aplikasi WHERE date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND userid='" & strUser_Selected & "'"
    Case 7 ' PTP
        strCaption = "PTP"
        sql_selected = "SELECT distinct y.agent,y.custid,x.name,y.promisedate FROM tblnegoptp_log y,mgm x WHERE y.custid=x.custid AND ( date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) AND (date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') ) AND y.agent='" & strUser_Selected & "' GROUP BY y.custid,y.agent,x.name,y.promisedate ORDER BY y.custid "
    Case 8 ' PTP REG
        strCaption = "PTP REG"
        sql_selected = "SELECT distinct y.agent,y.custid,x.name FROM tblnegoptp_log y,mgm x WHERE y.custid=x.custid AND date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND y.agent='" & strUser_Selected & "' GROUP BY y.custid,y.agent,x.name ORDER BY y.custid "
    Case 9 ' RPC1
        strCaption = "RPC 1"
        sql_selected = "SELECT custid,agent,f_cek_new,statuscall FROM mgm WHERE substring(f_cek_new,1,3) in('PTP','KP-','BP-','ON-','PR-') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND agent='" & strUser_Selected & "' "
    Case 10 ' RPC2
        strCaption = "RPC 2"
        sql_selected = "SELECT custid,agent,f_cek_new,statuscall FROM mgm WHERE trim(lower(statuscall)) in ('ch','spouse') AND agent like 'D%' AND date_part('month',tglcall)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglcall)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND agent='" & strUser_Selected & "'"
    Case 11 ' KEPT
        strCaption = "KEPT"
        sql_selected = "SELECT agent,x.custid,x.Tgl_janji,x.promisepay,y.paydate,y.payment FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY custid,agent,promisepay) x, " & _
                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y WHERE x.custid=y.custid AND (y.paydate between x.Tgl_janji AND x.Tgl_janji+5) AND y.payment>=x.promisepay AND agent='" & strUser_Selected & "'"
    Case 12 ' KEPT OLD
        strCaption = "KEPT OLD"
        sql_selected = "SELECT agent,custid,paydate,payment FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND agent='" & strUser_Selected & "' ORDER BY paydate,custid"
    Case 13 ' KEPT + BROKEN
        strCaption = "KEPT BROKEN"
        sql_selected = "SELECT x.agent,x.custid,x.Tgl_janji,x.promisepay,y.paydate,y.payment FROM (SELECT custid,agent,max(promisedate) as Tgl_janji,promisepay FROM tblnegoptp_log WHERE date_part('month',tglinput)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tglinput)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('month',promisedate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',promisedate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY custid,agent,promisepay) x LEFT JOIN " & _
                        " (SELECT custid,max(paydate) as paydate,payment FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-01") & "' AND '" & Format(DateAdd("m", 1, tgl_lap), "yyyy-mm-05") & "' GROUP BY custid,payment) y ON x.custid=y.custid WHERE (y.payment<x.promisepay OR y.payment is null) AND x.agent='" & strUser_Selected & "'"
    Case 14 ' KEPT + BROKEN OLD
        strCaption = "KEPT BROKEN OLD"
        sql_selected = "SELECT agent,custid,promisedate,promisepay FROM tblnegoptp_log WHERE custid not in ( SELECT distinct custid FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "')) AND promisedate between '" & Format(tgl_lap, "yyyy-mm-01") & "' and '" & Format(Now, "yyyy-mm-dd") & "' AND agent='" & strUser_Selected & "' ORDER BY promisedate "
    
    End Select
    
    If sql_selected <> "" Then
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open sql_selected
        If rs_temp.RecordCount > 0 Then
            With Form_detail_deskcoll
                .Caption = .Caption & " " & strCaption
                .show_data rs_temp
                .Show 1
            End With
        End If
    End If
End Sub


Public Sub ConvertToExcel_this(M_Objrs As ADODB.Recordset, Txtpath As String)
    Dim listItem        As listItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Double
    Dim m_msgbox As String
    Dim iCell           As Integer
    Dim iLastColumn     As Integer
    Dim arrAlpha
    
    i = 1
  
    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    arrAlpha = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD")
    
    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Double
        If M_Objrs.state = 1 Then
            x = 0
            Y = M_Objrs.fields().Count - 1
            Do Until x > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_Objrs.fields(x).Name)
                i = i + 1
                x = x + 1
            Loop
            objSheet.Cells(1, Y + 2).Value = "Dialer Op/Paid Hrs"
            objSheet.Cells(1, Y + 3).Value = "RPC/Dialer Op Hrs"
            objSheet.Cells(1, Y + 4).Value = "RPC/Paid Hrs"
            objSheet.Cells(1, Y + 5).Value = "(Kept+Broken)/RPC"
            objSheet.Cells(1, Y + 6).Value = "PTP/RPC"
            objSheet.Cells(1, Y + 7).Value = "Kept#/(Kept+Broken)"
            objSheet.Cells(1, Y + 8).Value = "Average Payment Size"
            objSheet.Cells(1, Y + 9).Value = "CEV"
            objSheet.Cells(1, Y + 10).Value = "EVPH"
            
            objSheet.Cells(1, Y + 11).Value = "(Kept+Broken)/RPC OLD"
            objSheet.Cells(1, Y + 12).Value = "PTP/RPC OLD"
            objSheet.Cells(1, Y + 13).Value = "Kept#/(Kept+Broken) OLD"
            objSheet.Cells(1, Y + 14).Value = "Average Payment Size OLD"
            objSheet.Cells(1, Y + 15).Value = "CEV OLD"
            objSheet.Cells(1, Y + 16).Value = "EVPH OLD"
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_Objrs '-> Proses pengisian data dimulai dari Cell A2
    M_Objrs.MoveFirst
    iCell = 2
    iLastColumn = M_Objrs.fields().Count
    While Not M_Objrs.EOF
'        objSheet.Cells(iCell, iLastColumn + 1).Value = "1"
        tPaid_hours = Val(cnull(rs_calc!paid_hours))
        tDialer_hours = Val(cnull(rs_calc!Dialer_hours))
        tPTP = Val(cnull(rs_calc!ptp))
        tRPC1 = Val(cnull(rs_calc!rpc1))
        tRPC2 = Val(cnull(rs_calc!rpc2))
        tKept = Val(cnull(rs_calc!kept))
        tKeptBroken = Val(cnull(rs_calc("Kept+Broken")))
        tKeptAmount = Val(cnull(rs_calc("Kept_amount")))
        tKept_old = Val(cnull(rs_calc!kept_old))
        tKeptBrokenOld = Val(cnull(rs_calc("Kept+Broken_old")))
        tPymnGlobal = Val(cnull(rs_calc("payment_global")))
        
        If tDialer_hours > 0 And tPaid_hours > 0 Then
            objSheet.Cells(iCell, iLastColumn + 1).Value = Round(tDialer_hours / tPaid_hours, 2)
        End If
        
        If tRPC2 > 0 And tDialer_hours > 0 Then
            objSheet.Cells(iCell, iLastColumn + 2).Value = Round(tRPC2 / tDialer_hours, 2)
        End If
        
        If tRPC2 > 0 And tPaid_hours > 0 Then
            objSheet.Cells(iCell, iLastColumn + 3).Value = Round(tRPC2 / tPaid_hours, 2)
        End If
        
        If tKeptBroken > 0 And tRPC2 > 0 Then
            objSheet.Cells(iCell, iLastColumn + 4).Value = Round(tKeptBroken / tRPC2, 2)
        End If
        
        If tPTP > 0 And tRPC2 > 0 Then
            objSheet.Cells(iCell, iLastColumn + 5).Value = Round(tPTP / tRPC2, 2)
        End If
        
        If tKept > 0 And tKeptBroken > 0 Then
            objSheet.Cells(iCell, iLastColumn + 6).Value = Round(tKept / tKeptBroken, 2)
        End If
        
        If tKeptAmount > 0 And tKept > 0 Then
            objSheet.Cells(iCell, iLastColumn + 7).Value = Round(tKeptAmount / tKept, 2)
        End If
        
        objSheet.Cells(iCell, iLastColumn + 8).Value = Val(objSheet.Cells(iCell, iLastColumn + 4).Value) * Val(objSheet.Cells(iCell, iLastColumn + 6).Value) * Val(objSheet.Cells(iCell, iLastColumn + 7).Value)
        objSheet.Cells(iCell, iLastColumn + 9).Value = Val(objSheet.Cells(iCell, iLastColumn + 3).Value) * Val(objSheet.Cells(iCell, iLastColumn + 8).Value)
        
        
        ' ---------- OLD CALCULATION ----------
        
        If tKeptBroken > 0 And tRPC2 > 0 Then
            objSheet.Cells(iCell, iLastColumn + 10).Value = Round(tKeptBrokenOld / tRPC2, 2)
        End If
        
        If tPTP > 0 And tRPC2 > 0 Then
            objSheet.Cells(iCell, iLastColumn + 11).Value = Round(tPTP / tRPC2, 2)
        End If
        
        If tKept > 0 And tKeptBroken > 0 Then
            objSheet.Cells(iCell, iLastColumn + 12).Value = Round(tKept_old / tKeptBrokenOld, 2)
        End If
        
        If tKeptAmount > 0 And tKept > 0 Then
            objSheet.Cells(iCell, iLastColumn + 13).Value = Round(tPymnGlobal / tKept_old, 2)
        End If
        
        objSheet.Cells(iCell, iLastColumn + 14).Value = Val(objSheet.Cells(iCell, iLastColumn + 10).Value) * Val(objSheet.Cells(iCell, iLastColumn + 12).Value) * Val(objSheet.Cells(iCell, iLastColumn + 13).Value)
        objSheet.Cells(iCell, iLastColumn + 15).Value = Val(objSheet.Cells(iCell, iLastColumn + 3).Value) * Val(objSheet.Cells(iCell, iLastColumn + 14).Value)
        
        ' --------------------------------------
        
        
        iCell = iCell + 1
        M_Objrs.MoveNext
    Wend
    objSheet.Cells(iCell, 1).Value = "TOTAL"
    
    
    
    For x = 4 To iLastColumn + 15
        objSheet.Cells(iCell, x).Value = "=sum(" & arrAlpha(x - 1) & "2:" & arrAlpha(x - 1) & iCell - 1 & ")"
    Next x
    
    objBook.SaveAs Txtpath, xlWorkbookNormal
    objExcel.Quit
    
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    'Set M_Objrs = Nothing
 
    Exit Sub
 
SALAH:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub msfx_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 ' If this is not row 0, do nothing.
    If msfx.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    SortByColumn msfx.MouseCol
End Sub

Private Sub SortByColumn(ByVal sort_column As Integer)
    ' Hide the FlexGrid.
    msfx.Visible = False
    msfx.Refresh

    ' Sort using the clicked column.
    msfx.Col = sort_column
    msfx.ColSel = sort_column
    msfx.Row = 0
    msfx.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    msfx.Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        msfx.TextMatrix(0, m_SortColumn) = Mid$(msfx.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        msfx.TextMatrix(0, m_SortColumn) = "> " & msfx.TextMatrix(0, m_SortColumn)
    Else
        msfx.TextMatrix(0, m_SortColumn) = "< " & msfx.TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    msfx.Visible = True
End Sub
