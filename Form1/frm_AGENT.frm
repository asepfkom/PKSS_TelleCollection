VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_AGENT 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7590
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5040
      TabIndex        =   22
      Top             =   1440
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4560
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   20
      Top             =   2865
      Width           =   2385
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1215
      Width           =   2370
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      TabIndex        =   7
      Top             =   2190
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Resign"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   1
      Left            =   2565
      TabIndex        =   9
      Top             =   2550
      Width           =   1290
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000004&
      Caption         =   "Works"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Index           =   0
      Left            =   1545
      TabIndex        =   8
      Top             =   2550
      Width           =   1125
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1545
      MaxLength       =   13
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1545
      Width           =   2370
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   3330
      TabIndex        =   3
      Top             =   885
      Width           =   4140
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   1545
      TabIndex        =   2
      Top             =   885
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   6285
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   5325
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      MaxLength       =   50
      TabIndex        =   1
      Top             =   555
      Width           =   5925
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1545
      MaxLength       =   20
      TabIndex        =   0
      Top             =   225
      Width           =   2370
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   315
      Left            =   1545
      TabIndex        =   6
      Top             =   1875
      Visible         =   0   'False
      Width           =   2280
      _Version        =   65536
      _ExtentX        =   4022
      _ExtentY        =   556
      Calculator      =   "frm_AGENT.frx":0000
      Caption         =   "frm_AGENT.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_AGENT.frx":008C
      Keys            =   "frm_AGENT.frx":00AA
      Spin            =   "frm_AGENT.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   0
      Format          =   "###,###,##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   9999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "AM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2895
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   6
      Left            =   135
      TabIndex        =   19
      Top             =   1260
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   5
      Left            =   135
      TabIndex        =   18
      Top             =   2220
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Status Agent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   17
      Top             =   2565
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Team"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   135
      TabIndex        =   16
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Basic Salary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   15
      Top             =   1905
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Nama Supervisor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   135
      TabIndex        =   14
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Nama Agent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Index           =   0
      Left            =   135
      TabIndex        =   13
      Top             =   585
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "User Id"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   150
      TabIndex        =   12
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frm_AGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLSSPV_AGENT
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = M_DATA.COMBO_SPV(M_OBJCONN, " spvcode =  '" + Combo1(0).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("spvcode")
        Combo1(1).Text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        Text4.Text = IIf(IsNull(M_objrs("TEAM")), "", M_objrs("TEAM"))
        Text5.Text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Text4.Text = Empty
        Text5.Text = Empty
    End If
Case 1
    Set M_objrs = M_DATA.COMBO_SPV(M_OBJCONN, " agent =  '" + Combo1(1).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("spvcode")
        Combo1(1).Text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        Text4.Text = IIf(IsNull(M_objrs("TEAM")), "", M_objrs("TEAM"))
        Text5.Text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Text4.Text = Empty
        Text5.Text = Empty
    End If
End Select
Set M_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLSSPV_AGENT
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_objrs = M_DATA.COMBO_SPV(M_OBJCONN, " spvcode =  '" + Combo1(0).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("SPVCODE")
        Combo1(1).Text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        Text4.Text = IIf(IsNull(M_objrs("TEAM")), "", M_objrs("TEAM"))
        Text5.Text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Text4.Text = Empty
        Text5.Text = Empty
    End If
Case 1
    Set M_objrs = M_DATA.COMBO_SPV(M_OBJCONN, " agent =  '" + Combo1(1).Text + "'")
    If M_objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_objrs("SPVCODE")
        Combo1(1).Text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        Text4.Text = IIf(IsNull(M_objrs("TEAM")), "", M_objrs("TEAM"))
        Text5.Text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Text4.Text = Empty
        Text5.Text = Empty
    End If
End Select
Set M_objrs = Nothing
Set M_DATA = Nothing
End Sub
Private Sub Combo2_LostFocus()
Select Case UCase(Combo2.Text)
    Case "SENIOR"
    Case "JUNIOR"
    Case "TRAINEE"
    Case Else
        Combo2.Text = Empty
End Select
End Sub
Private Sub Command1_Click(Index As Integer)
Dim VSAVE As Boolean
Dim listdo As String
Dim sql As String
VSAVE = True
Select Case Index
    Case 0
        VSAVE = VSAVE And Text1.Text <> Empty
        VSAVE = VSAVE And Text2.Text <> Empty
        VSAVE = VSAVE And Combo1(0).Text <> Empty
        VSAVE = VSAVE And Combo2.Text <> Empty
        If VSAVE Then
            ok = True
            Me.Hide
            'FRM_AGENT_LIST.ListView1.SetFocus
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
            If Text6.Text = "" Then
                listdo = "ADD"
                sql = "INSERT INTO hst_telecollection VALUES (now(),'" + Text2.Text + "','','" + MDIForm1.txtlevel.Text + "','" + listdo + "')"
            ElseIf Text6.Text <> "" Then
                listdo = "EDIT"
                sql = "INSERT INTO hst_telecollection VALUES (now(),'" + Text6.Text + "','" + Text2.Text + "','" + MDIForm1.txtlevel.Text + "','" + listdo + "')"
            End If
            M_OBJCONN.Execute sql
    Case 1
        ok = False
        Unload Me
        'FRM_AGENT_LIST.ListView1.SetFocus
End Select
End Sub

Private Sub Form_Load()
Dim M_objrs As ADODB.Recordset
Dim M_DATA As New CLSSPV_AGENT
Set M_objrs = M_DATA.COMBO_SPV(M_OBJCONN, "")
    While Not M_objrs.EOF
        Combo1(0).AddItem M_objrs("spvcode")
        Combo1(0).DataField = M_objrs("spvcode")
        Combo1(1).AddItem M_objrs("agent")
        Combo1(1).DataField = M_objrs("agent")
        M_objrs.MoveNext
    Wend
        Combo2.AddItem "Senior", 0
        Combo2.AddItem "Junior", 1
        Combo2.AddItem "Trainee", 2
    TDBNumber1.Value = 0
Set M_objrs = Nothing
Set M_DATA = Nothing
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click(0)
End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hwnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).list(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub
