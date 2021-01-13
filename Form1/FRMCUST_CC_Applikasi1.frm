VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form FRMCUST_CC_Applikasi1 
   Caption         =   "Form2"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   1755
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   600
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   556
      Calendar        =   "FRMCUST_CC_Applikasi1.frx":0000
      Caption         =   "FRMCUST_CC_Applikasi1.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRMCUST_CC_Applikasi1.frx":0184
      Keys            =   "FRMCUST_CC_Applikasi1.frx":01A2
      Spin            =   "FRMCUST_CC_Applikasi1.frx":0200
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
      Text            =   "11/05/2004"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   38118
      CenturyMode     =   0
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   270
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   1140
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   390
      Index           =   0
      Left            =   2370
      TabIndex        =   2
      Top             =   1140
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Of Birth :"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Add On Name :"
      Height          =   225
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   300
      Width           =   1260
   End
End
Attribute VB_Name = "FRMCUST_CC_Applikasi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ok As Boolean

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            ok = True
            Me.Hide
        Case 1
            ok = False
            Unload Me
    End Select
End Sub

