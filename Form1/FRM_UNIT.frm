VERSION 5.00
Begin VB.Form FRM_UNIT 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   360
      Index           =   1
      Left            =   3945
      TabIndex        =   3
      Top             =   630
      Width           =   1035
   End
   Begin VB.ComboBox Combo1 
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   2100
      TabIndex        =   1
      Top             =   135
      Width           =   2100
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Tampilkan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2790
      TabIndex        =   0
      Top             =   615
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Unit Penjualan :"
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
      Left            =   525
      TabIndex        =   2
      Top             =   165
      Width           =   1530
   End
End
Attribute VB_Name = "FRM_UNIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0
If KeyAscii = 13 Then
   Combo1_LostFocus (Index)
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
Select Case Index
    Case 0
        Select Case UCase(Combo1(0).Text)
        Case "KTA", "KTA - CROSS SELL", "CREDIT CARD", "CC - CROSS SELL"
        Case ""
        Case Else
            MsgBox "Pilih Salah Satu Dari Pilihan Yang Tersedia"
            Combo1(0).Text = Empty
            Exit Sub
        End Select
End Select
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        UNIT = Combo1(0).Text
        TL = ""
        FRM_LOGIN_LOGOUT.Show vbModal
    Case 1
        Unload Me
End Select
End Sub

Private Sub Form_Load()
    Combo1(0).AddItem "KTA", 0
    Combo1(0).AddItem "KTA - CROSS SELL", 1
    Combo1(0).AddItem "CC - CROSS SELL", 1
    Combo1(0).AddItem "Credit Card", 2
End Sub
