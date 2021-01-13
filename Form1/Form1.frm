VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBTime6Ctl.TDBTime TDBTime1 
      Height          =   315
      Left            =   2835
      TabIndex        =   9
      Top             =   570
      Width           =   1080
      _Version        =   65536
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "Form1.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "Form1.frx":0936
      Spin            =   "Form1.frx":0986
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn:ss"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   0
      Format          =   "hh:nn:ss"
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
      Text            =   "09:58:56"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0.415925925925926
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1395
      TabIndex        =   8
      Top             =   570
      Width           =   1425
      _Version        =   65536
      _ExtentX        =   2514
      _ExtentY        =   556
      Calendar        =   "Form1.frx":09AE
      Caption         =   "Form1.frx":0AC6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form1.frx":0B32
      Keys            =   "Form1.frx":0B50
      Spin            =   "Form1.frx":0BAE
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
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
      ForeColor       =   0
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
      Text            =   "11/11/2002"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   37571
      CenturyMode     =   0
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
      Height          =   375
      Index           =   0
      Left            =   2850
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1245
      UseMaskColor    =   -1  'True
      Width           =   825
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
      Height          =   375
      Index           =   1
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1245
      UseMaskColor    =   -1  'True
      Width           =   810
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2835
      MaxLength       =   12
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1395
      MaxLength       =   20
      TabIndex        =   4
      Top             =   135
      Width           =   1605
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1395
      MaxLength       =   20
      TabIndex        =   1
      Top             =   885
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tanggal"
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   615
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Id"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   480
      TabIndex        =   6
      Top             =   210
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Jumlah"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   0
      Left            =   495
      TabIndex        =   5
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        ok = True
        If Text1.Text = Empty Or Text2.Text = Empty Or TDBDate1.ValueIsNull Or TDBTime1.ValueIsNull Then
            MsgBox "Data Tidak Lengkap"
            Exit Sub
        End If
        
        Me.Hide
    Case 1
        ok = False
        Unload Me
End Select
End Sub

Private Sub Text2_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub Text3_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub
