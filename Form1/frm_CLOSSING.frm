VERSION 5.00
Begin VB.Form frm_CLOSSING 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   900
      Width           =   1500
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
      Height          =   390
      Index           =   1
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
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
      Height          =   390
      Index           =   0
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
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
      MaxLength       =   2
      TabIndex        =   0
      Top             =   225
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Untuk Data"
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
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Keterangan"
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
      Height          =   375
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   585
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Caption         =   "Kode"
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
      Height          =   375
      Left            =   135
      TabIndex        =   5
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "frm_CLOSSING"
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
            FRM_CLOSSING_LIST.ListView1.SetFocus
        Else
            MsgBox "Data Yang Anda Masukan Tidak Lengkap", vbInformation, "Informasi"
        End If
    Case 1
        ok = False
        Unload Me
        FRM_CLOSSING_LIST.ListView1.SetFocus
End Select
End Sub

Private Sub Form_Load()
    Combo1.AddItem "LEADS"
    Combo1.AddItem "mgm"
    Combo1.AddItem "BOTH"
End Sub
