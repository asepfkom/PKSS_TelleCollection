VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FRM_SCHEDULE 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "FRM_SCHEDULE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRM_SCHEDULE.frx":08CA
   ScaleHeight     =   2400
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   975
      Width           =   225
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1080
      Left            =   990
      TabIndex        =   10
      Top             =   1260
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   1905
      _Version        =   196610
      BackColor       =   -2147483644
      BackStyle       =   1
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   75
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "FRM_SCHEDULE.frx":267D3
         Caption         =   "FRM_SCHEDULE.frx":268EB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FRM_SCHEDULE.frx":26957
         Keys            =   "FRM_SCHEDULE.frx":26975
         Spin            =   "FRM_SCHEDULE.frx":269D3
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
         ForeColor       =   -2147483640
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
         Value           =   37609
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Index           =   1
         Left            =   1830
         TabIndex        =   2
         Top             =   90
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "FRM_SCHEDULE.frx":269FB
         Caption         =   "FRM_SCHEDULE.frx":26B13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FRM_SCHEDULE.frx":26B7F
         Keys            =   "FRM_SCHEDULE.frx":26B9D
         Spin            =   "FRM_SCHEDULE.frx":26BFB
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
         ForeColor       =   -2147483640
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
         Value           =   37609
         CenturyMode     =   0
      End
      Begin Threed.SSCommand CmdScheduleFInd 
         Height          =   465
         Index           =   3
         Left            =   90
         TabIndex        =   3
         Top             =   450
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   820
         _Version        =   196610
         Font3D          =   4
         MousePointer    =   16
         ForeColor       =   0
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Show"
         ButtonStyle     =   2
         BevelWidth      =   3
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "S/d"
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   0
         Left            =   1485
         TabIndex        =   11
         Top             =   90
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proses....!!"
      Height          =   720
      Left            =   210
      TabIndex        =   8
      Top             =   4110
      Visible         =   0   'False
      Width           =   3990
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   390
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
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
      Index           =   1
      Left            =   1425
      TabIndex        =   0
      Top             =   1095
      Width           =   3480
   End
   Begin Threed.SSCommand CmdScheduleClear 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   4530
      TabIndex        =   4
      Top             =   1935
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
      Caption         =   "&Cancel"
      ButtonStyle     =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   1085
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
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
      Caption         =   "Schedule"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   1470
      TabIndex        =   6
      Top             =   690
      Visible         =   0   'False
      Width           =   1905
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
      Left            =   2310
      TabIndex        =   13
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Name :"
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
      Height          =   300
      Index           =   1
      Left            =   225
      TabIndex        =   7
      Top             =   735
      Width           =   1170
   End
End
Attribute VB_Name = "FRM_SCHEDULE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check1.Value = 1 Then
    StsMgmSchedule = True
Else
    StsMgmSchedule = False
End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 1
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
Dim m_data As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
End Select
Set m_data = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim m_data As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
End Select
Set m_data = Nothing
Set m_objrs = Nothing
End Sub


Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim m_data As New CLS_FRMSEARCH

Set m_objrs = m_data.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("USERID")
        Combo1(0).DataField = m_objrs("USERID")
        Combo1(1).AddItem m_objrs("AGENT")
        Combo1(1).DataField = m_objrs("AGENT")
        m_objrs.MoveNext
    Wend
Set m_data = Nothing
Set m_objrs = Nothing

Me.Top = 1500
Me.Left = 5000
StsMgmSchedule = False

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    Combo1(0).Text = MDIForm1.Text1.Text
    Combo1(0).Visible = False
    Combo1(1).Visible = False
    Label1(1).Visible = False
Else
    Combo1(0).Visible = False
    Combo1(1).Visible = True
    Label1(1).Visible = True
End If
End Sub

Private Sub Form_Terminate()
    ProgressBar1.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProgressBar1.Value = 0
End Sub

Private Sub SSCommand1_Click(Index As Integer)
If Combo1(0).Text = Empty Then
    MsgBox "Agent Harus Diisi", vbCritical + vbOKOnly, "Informasi"
    Exit Sub
End If
Select Case Index
Case 3
    If TDBDate1(0).ValueIsNull Or TDBDate1(1).ValueIsNull Then
        MsgBox "Tanggal Tidak Boleh Kosong", vbInformation + vbOKOnly, "Informasi"
        Exit Sub
    End If
    If TDBDate1(0).Value > TDBDate1(1).Value Then
        MsgBox "Tanggal Periode Awal harus Lebih Kecil Dari Tanggal Periode Akhir", vbInformation + vbOKOnly, "Informasi"
        Exit Sub
    End If
    search_ok = False
    If StsMgmSchedule = True Then
        FRM_PRESCREEN.Caption = "Data Mgm Schedule dari Tgl " & TDBDate1(0).Text & " Sampai Dengan " & TDBDate1(1).Text
    Else
        FRM_PRESCREEN.Caption = "Data Referall dari Tgl " & TDBDate1(0).Text & " Sampai Dengan " & TDBDate1(1).Text
    End If
    
    FRM_PRESCREEN.Show
End Select
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub
