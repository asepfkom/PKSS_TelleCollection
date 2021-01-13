VERSION 5.00
Begin VB.Form FBILL 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "BILLING"
   ClientHeight    =   1455
   ClientLeft      =   11355
   ClientTop       =   210
   ClientWidth     =   4320
   LinkTopic       =   "Form2"
   ScaleHeight     =   1455
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   120
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   2355
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING READY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label11 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1200
      TabIndex        =   1
      Top             =   450
      Width           =   225
   End
   Begin VB.Label Label12 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   780
      TabIndex        =   0
      Top             =   450
      Width           =   375
   End
End
Attribute VB_Name = "FBILL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
MDIForm1.ActionCTI ("HANGUP")
Call savecall
FBILL.Timer6.Enabled = False
Unload FBILL
End Sub

Private Sub Form_Load()
detik1 = 0
menit1 = 0
jam1 = 0
Label12.Caption = 0
Label10.Caption = 0
Label8.Caption = 0
Cnt = 0
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'totcost = 0
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Timer6_Timer()
If detik1 < 59 Then
    detik1 = detik1 + 1
Else
    detik1 = 0
    If menit1 < 59 Then
        menit1 = menit1 + 1
    Else
        menit1 = 0
        If jam1 < 23 Then
            jam1 = jam1 + 1
        Else
            detik1 = 0
            menit1 = 0
            jam1 = 0
        End If
    End If
End If
Cnt = Cnt + 1
If Cnt Mod rounding = 0 Then
    totcost = totcost + tarif
End If
Label8.Caption = detik1
Label10.Caption = menit1
Label12.Caption = jam1
Label3.Caption = "Cost : " & totcost
End Sub
