VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9090
   ControlBox      =   0   'False
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5895
      TabIndex        =   1
      Top             =   5415
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   397
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
      Left            =   5685
      TabIndex        =   0
      Top             =   5670
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim i As Double
Dim a As Integer
ProgressBar1.Max = 6000
ProgressBar1.Visible = True
For a = 0 To 3
    ProgressBar1.Value = 0
    For i = 0 To 6000
        ProgressBar1.Value = i
    Next i
Next a
    Unload Me
End Sub
