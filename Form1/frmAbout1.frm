VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5160
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9465
   ControlBox      =   0   'False
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   6525
      TabIndex        =   1
      Top             =   4365
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   240
      Top             =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Application System ....."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6165
      TabIndex        =   0
      Top             =   4725
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim i As Double
Dim a As Integer
ProgressBar1.Max = 6000
ProgressBar1.Visible = True
'For a = 0 To 3
    ProgressBar1.Value = 0
    For i = 0 To 6000
        ProgressBar1.Value = i
    Next i
'Next a
    Unload Me
    frmLogin.Show
End Sub
