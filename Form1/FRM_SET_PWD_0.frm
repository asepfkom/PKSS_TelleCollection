VERSION 5.00
Begin VB.Form FRM_SET_PWD_0 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1140
      MaxLength       =   6
      TabIndex        =   6
      Top             =   795
      Width           =   1230
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
      Height          =   585
      Index           =   1
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1215
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
      Height          =   585
      Index           =   0
      Left            =   5265
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1215
      UseMaskColor    =   -1  'True
      Width           =   825
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
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   1
      Top             =   165
      Width           =   1755
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
      Left            =   1140
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   5925
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Login PABX"
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
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "User Name"
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
      Height          =   345
      Left            =   45
      TabIndex        =   3
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Password"
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
      Left            =   45
      TabIndex        =   2
      Top             =   525
      Width           =   1035
   End
End
Attribute VB_Name = "FRM_SET_PWD_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ok As Boolean

Private Sub Command1_Click(Index As Integer)
Dim VSAVE As Boolean
VSAVE = True
Select Case Index
    Case 0
        VSAVE = VSAVE And Text1.Text <> Empty
        VSAVE = VSAVE And Text2.Text <> Empty
        If VSAVE Then
            ok = True
            Me.Hide
            FRM_SET_PWD.ListView1.SetFocus
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
    Case 1
        ok = False
        Unload Me
        FRM_SET_PWD.ListView1.SetFocus
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 32
    KeyAscii = 0
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
    Case Else
        KeyAscii = 0
End Select
End Sub
