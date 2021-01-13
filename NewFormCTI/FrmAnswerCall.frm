VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmAnswerCall 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1770
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3060
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1125
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   1984
      _Version        =   196610
      Font3D          =   2
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   2490
         Top             =   60
      End
      Begin Threed.SSCommand CmdAnswer 
         Height          =   405
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         _Version        =   196610
         Font3D          =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Accept"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand CmdAnswer 
         Height          =   405
         Index           =   1
         Left            =   1545
         TabIndex        =   2
         Top             =   645
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   714
         _Version        =   196610
         Font3D          =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Close"
         ButtonStyle     =   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "..Ringing.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   135
         TabIndex        =   3
         Top             =   30
         Width           =   2580
      End
   End
End
Attribute VB_Name = "FrmAnswerCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnswer_Click(Index As Integer)
Select Case Index
Case 0
    Call MDIForm1.ActionCTI("ACCEPT")
    
Case 1
    Call MDIForm1.ActionCTI("REJECT")
    Unload Me
End Select
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = False Then
    Label1.Visible = True
Else
    Label1.Visible = False
End If
End Sub


