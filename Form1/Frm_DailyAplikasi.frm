VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Frm_DailyAplikasi 
   Caption         =   "Daily Application"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   Icon            =   "Frm_DailyAplikasi.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8430
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Maintanance"
      Height          =   2160
      Left            =   5565
      TabIndex        =   11
      Top             =   90
      Width           =   5805
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Add"
         Height          =   330
         Index           =   0
         Left            =   2520
         TabIndex        =   26
         Top             =   1590
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   330
         Index           =   1
         Left            =   3300
         TabIndex        =   25
         Top             =   1590
         Width           =   795
      End
      Begin VB.TextBox TxtcNamaAppl 
         Height          =   315
         Left            =   1680
         TabIndex        =   18
         Top             =   195
         Width           =   3195
      End
      Begin VB.TextBox TxtcNoTelp 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   510
         Width           =   3195
      End
      Begin VB.ComboBox CmbcSalesForceId 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox CmbcSalesForceId 
         Height          =   315
         Index           =   1
         Left            =   2790
         TabIndex        =   15
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox CmbcJnsAppl 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   1170
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Delete"
         Height          =   330
         Index           =   2
         Left            =   4095
         TabIndex        =   13
         Top             =   1590
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   330
         Index           =   3
         Left            =   4905
         TabIndex        =   12
         Top             =   1575
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama Di Applikasi :"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "No Telp :"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   540
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Force Id :"
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   20
         Top             =   870
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Jenis Applikasi :"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "List Data"
      Height          =   2085
      Left            =   45
      TabIndex        =   0
      Top             =   150
      Width           =   5400
      Begin VB.CommandButton Command3 
         Caption         =   "&Reset"
         Height          =   300
         Left            =   4005
         TabIndex        =   24
         Top             =   1590
         Width           =   990
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Search"
         Height          =   300
         Left            =   2880
         TabIndex        =   23
         Top             =   1590
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1260
         TabIndex        =   9
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Top             =   825
         Width           =   1920
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1260
         TabIndex        =   6
         Top             =   540
         Width           =   3615
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   285
         Left            =   1260
         TabIndex        =   1
         Top             =   240
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   503
         Calendar        =   "Frm_DailyAplikasi.frx":000C
         Caption         =   "Frm_DailyAplikasi.frx":0124
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_DailyAplikasi.frx":0190
         Keys            =   "Frm_DailyAplikasi.frx":01AE
         Spin            =   "Frm_DailyAplikasi.frx":020C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.44024378152593E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate2 
         Height          =   285
         Left            =   2970
         TabIndex        =   2
         Top             =   240
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   503
         Calendar        =   "Frm_DailyAplikasi.frx":0234
         Caption         =   "Frm_DailyAplikasi.frx":034C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Frm_DailyAplikasi.frx":03B8
         Keys            =   "Frm_DailyAplikasi.frx":03D6
         Spin            =   "Frm_DailyAplikasi.frx":0434
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         Value           =   2.44024378152593E-316
         CenturyMode     =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Jenis Applikasi :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   1140
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sales Force Id :"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   870
         Width           =   1110
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama :"
         Height          =   240
         Left            =   510
         TabIndex        =   5
         Top             =   585
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "s/d"
         Height          =   240
         Left            =   2625
         TabIndex        =   4
         Top             =   255
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tgl Input :"
         Height          =   240
         Left            =   495
         TabIndex        =   3
         Top             =   270
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6060
      Left            =   75
      TabIndex        =   27
      Top             =   2310
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   10689
      View            =   3
      LabelEdit       =   1
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
      Picture         =   "Frm_DailyAplikasi.frx":045C
   End
End
Attribute VB_Name = "Frm_DailyAplikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
        
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
Dim m_user As New ADODB.Recordset
Dim listitem As listitem
On Error GoTo LoadErr:
Call header
m_user.CursorLocation = adUseClient
m_user.Open "Select * from usertbl where usertype =1 order by userid", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not m_user.EOF
    CmbcSalesForceId(0).AddItem m_user!USERID
    CmbcSalesForceId(1).AddItem m_user!agent
    m_user.MoveNext
Wend
Set m_user = Nothing
Exit Sub
LoadErr:
Set m_user = Nothing
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "No", 10 * 120
    ListView1.ColumnHeaders.ADD 3, , "Tgl Input", 15 * 120
    ListView1.ColumnHeaders.ADD 4, , "Sales Force Id", 15 * 120
    ListView1.ColumnHeaders.ADD 5, , "No Telp", 15 * 120
    ListView1.ColumnHeaders.ADD 6, , "Nama Di Aplikasi", 30 * 120
    ListView1.ColumnHeaders.ADD 7, , "No Telp", 15 * 120
End Sub

